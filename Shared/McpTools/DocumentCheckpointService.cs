using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class DocumentCheckpointService
    {
        private const int MaxCheckpointSlots = 5;

        private static readonly Dictionary<string, DocumentCheckpointSession> Sessions =
            new Dictionary<string, DocumentCheckpointSession>(StringComparer.Ordinal);

        private static Word.Application _wordApplication;

        public static void SetWordApplication(object wordApplication)
        {
            _wordApplication = wordApplication as Word.Application;
        }

        /// <summary>
        /// 从文档 RCW 回填 Application（Desktop + WPS 兼容路径常无宿主注入）。
        /// </summary>
        private static void TryCacheApplicationFromDocument(Word.Document document)
        {
            if (document == null || _wordApplication != null)
            {
                return;
            }

            try
            {
                Word.Application app = document.Application;
                if (app != null)
                {
                    _wordApplication = app;
                }
            }
            catch (Exception)
            {
            }
        }

        public static string GetCheckpointRootDirectory()
        {
            string root = McpToolsHelpers.GetUserWorkDirectory("document_checkpoints");
            if (string.IsNullOrEmpty(root))
            {
                root = Path.Combine(Path.GetTempPath(), "EasyWrite", "document_checkpoints");
            }

            Directory.CreateDirectory(root);
            return root;
        }

        public static string GetDocCheckpointDirectory(string docUuid)
        {
            return Path.Combine(GetCheckpointRootDirectory(), docUuid ?? "");
        }

        public static string GetCheckpointFilePath(string docUuid, string timestamp)
        {
            return Path.Combine(GetDocCheckpointDirectory(docUuid), $"checkpoint_{timestamp}.docx");
        }

        public static DocumentCheckpointSession GetSession(string docUuid)
        {
            if (string.IsNullOrEmpty(docUuid))
            {
                return null;
            }

            Sessions.TryGetValue(docUuid, out DocumentCheckpointSession session);
            return session;
        }

        private static DocumentCheckpointSession GetOrCreateSession(string docUuid)
        {
            if (string.IsNullOrEmpty(docUuid))
            {
                throw new ArgumentException("docUuid is required", nameof(docUuid));
            }

            if (!Sessions.TryGetValue(docUuid, out DocumentCheckpointSession session))
            {
                session = new DocumentCheckpointSession { DocUuid = docUuid };
                Sessions[docUuid] = session;
            }

            return session;
        }

        public static CheckpointBeforeResult CreateBeforeMutation(
            string toolName,
            IReadOnlyDictionary<string, object> parameters)
        {
            // 优先按 channel_id / 默认渠道解析（Desktop+WPS 无 Word.Application 注入）
            Word.Document activeDoc = null;
            string resolveHint = null;
            Dictionary<string, object> args = ToArgsDictionary(parameters);
            if (DocumentHostAdapter.TryResolveInteropDocument(
                    args,
                    _wordApplication,
                    out InteropDocumentHandle handle,
                    out ToolResult resolveError))
            {
                activeDoc = handle.Document;
                TryCacheApplicationFromDocument(activeDoc);
                System.Diagnostics.Debug.WriteLine(
                    $"[Checkpoint] resolve via DocumentHost host={handle.HostName} channel_id={handle.ChannelId} tool={toolName}");
            }
            else
            {
                resolveHint = resolveError?.Error;
            }

            // 回退：宿主已注入 Word.Application（Plugin / F_open_word 后）
            if (activeDoc == null && _wordApplication != null)
            {
                try
                {
                    activeDoc = _wordApplication.ActiveDocument;
                }
                catch (Exception ex)
                {
                    return Fail("无法读取活动文档: " + ex.Message);
                }
            }

            if (activeDoc == null)
            {
                string detail = string.IsNullOrEmpty(resolveHint)
                    ? "无可用文档渠道；请确认 WPS/Word 已打开文档，或传入 channel_id"
                    : resolveHint;
                return Fail(detail);
            }

            DocumentState.BindAndActivate(activeDoc);
            string docUuid = DocumentState.EnsureCurrDocUuid(activeDoc);

            string dir = GetDocCheckpointDirectory(docUuid);
            Directory.CreateDirectory(dir);

            string timestamp = GenerateUniqueTimestamp(dir);
            string path = GetCheckpointFilePath(docUuid, timestamp);

            try
            {
                DocumentCheckpointSaveHelper.SaveDocumentCopy(activeDoc, path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] Save document copy failed: {ex.Message}");
                return Fail($"创建文档副本失败: {ex.Message}");
            }

            DocumentCheckpointSession session = GetOrCreateSession(docUuid);
            AppendCheckpointEntry(session, timestamp);

            if (session.CheckpointTimestampsInOrder.Count > MaxCheckpointSlots)
            {
                EvictOldestCheckpoint(session, docUuid);
            }

            System.Diagnostics.Debug.WriteLine(
                $"[Checkpoint] copy saved path={path} timestamp={timestamp} historyCount={session.CheckpointTimestampsInOrder.Count}");

            return new CheckpointBeforeResult
            {
                Success = true,
                Timestamp = timestamp,
                DocUuid = docUuid
            };
        }

        private static Dictionary<string, object> ToArgsDictionary(IReadOnlyDictionary<string, object> parameters)
        {
            if (parameters == null)
            {
                return new Dictionary<string, object>();
            }

            if (parameters is Dictionary<string, object> dict)
            {
                return dict;
            }

            var copy = new Dictionary<string, object>();
            foreach (var kv in parameters)
            {
                copy[kv.Key] = kv.Value;
            }

            return copy;
        }

        public static void AppendActionLogAfterMutation(
            string toolName,
            IReadOnlyDictionary<string, object> parameters)
        {
            string docUuid = DocumentState.CurrDocUuid;
            if (string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            DocumentCheckpointSession session = GetOrCreateSession(docUuid);

            if (string.Equals(toolName, "F_process_document_actions", StringComparison.Ordinal))
            {
                AppendProcessActionsLogs(session, toolName, parameters);
            }
            else
            {
                string json = OperationLogFormatter.FormatArgumentsJson(parameters);
                session.Entries.Add(new OperationLogEntry
                {
                    Kind = OperationLogEntryKind.Action,
                    ToolName = toolName,
                    ActionLineBody = json
                });
            }

            System.Diagnostics.Debug.WriteLine($"[Checkpoint] append action tool={toolName}");
        }

        public static bool IsInvokeMutationFailed(ToolResult result, string toolName)
        {
            if (result == null || !result.Success)
            {
                return true;
            }

            if (!string.Equals(toolName, "F_process_document_actions", StringComparison.Ordinal))
            {
                return false;
            }

            int? total = GetIntProperty(result.Data, "total_actions");
            int? executed = GetIntProperty(result.Data, "executed_actions");
            if (total.HasValue && executed.HasValue && executed.Value < total.Value)
            {
                return true;
            }

            object resultsObj = GetPropertyValue(result.Data, "results");
            return resultsObj != null && HasAnyFailedActionResult(resultsObj);
        }

        public static bool IsCheckpointCleanupOnly(ToolResult result)
        {
            object value = GetPropertyValue(result?.Data, "checkpoint_cleanup_only");
            if (value is bool boolValue)
            {
                return boolValue;
            }

            if (value != null && bool.TryParse(value.ToString(), out bool parsed))
            {
                return parsed;
            }

            return false;
        }

        public static void CleanupFailedInvoke(string docUuid, string timestamp)
        {
            if (string.IsNullOrEmpty(docUuid) || string.IsNullOrEmpty(timestamp))
            {
                return;
            }

            RemoveCheckpointAndFollowingActions(docUuid, timestamp, deleteNewerFiles: false);
            DeleteCheckpointFile(docUuid, timestamp);
            System.Diagnostics.Debug.WriteLine($"[Checkpoint] cleanup-only failed invoke timestamp={timestamp}");
        }

        private static int? GetIntProperty(object data, string name)
        {
            object value = GetPropertyValue(data, name);
            if (value == null)
            {
                return null;
            }

            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return null;
            }
        }

        private static object GetPropertyValue(object data, string name)
        {
            if (data == null)
            {
                return null;
            }

            if (data is Dictionary<string, object> dict && dict.TryGetValue(name, out object dictValue))
            {
                return dictValue;
            }

            if (data is IDictionary legacy && legacy.Contains(name))
            {
                return legacy[name];
            }

            PropertyInfo prop = data.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            return prop?.GetValue(data);
        }

        public static void RollbackFailedMutation(string docUuid, string timestamp)
        {
            if (string.IsNullOrEmpty(docUuid) || string.IsNullOrEmpty(timestamp))
            {
                return;
            }

            string path = GetCheckpointFilePath(docUuid, timestamp);
            if (File.Exists(path) && _wordApplication?.ActiveDocument != null)
            {
                try
                {
                    DocumentRestoreHelper.RestoreActiveDocumentFromCheckpoint(path, _wordApplication);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Checkpoint] rollback restore warning: {ex.Message}");
                }
            }

            DocumentState.ClearAfterCheckpointRestore();
            RemoveCheckpointAndFollowingActions(docUuid, timestamp, deleteNewerFiles: false);
            DeleteCheckpointFile(docUuid, timestamp);

            System.Diagnostics.Debug.WriteLine($"[Checkpoint] rollback failed invoke timestamp={timestamp}");
        }

        public static ToolResult RestoreCheckpoint(string timestamp)
        {
            if (_wordApplication?.ActiveDocument == null)
            {
                return new ToolResult { Success = false, Error = "没有活动的 Word 文档" };
            }

            Word.Document activeDoc = _wordApplication.ActiveDocument;
            DocumentState.BindAndActivate(activeDoc);
            string docUuid = DocumentState.CurrDocUuid;
            if (string.IsNullOrEmpty(docUuid))
            {
                return new ToolResult { Success = false, Error = "当前文档尚未绑定 doc_uuid" };
            }

            DocumentCheckpointSession session = GetSession(docUuid);
            if (session == null || !session.CheckpointTimestampsInOrder.Contains(timestamp))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = $"副本 timestamp '{timestamp}' 不存在或已被淘汰。请先调用 F_list_document_operations 查看可用副本。"
                };
            }

            string path = GetCheckpointFilePath(docUuid, timestamp);
            if (!File.Exists(path))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = $"副本文件不存在: checkpoint_{timestamp}.docx。请先调用 F_list_document_operations。"
                };
            }

            try
            {
                DocumentRestoreHelper.RestoreActiveDocumentFromCheckpoint(path, _wordApplication);
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = $"恢复副本失败: {ex.Message}" };
            }

            DocumentState.ClearAfterCheckpointRestore();
            int removed = TruncateLogAfterCheckpoint(docUuid, timestamp);

            System.Diagnostics.Debug.WriteLine(
                $"[Checkpoint] restore timestamp={timestamp} truncate removed={removed}");

            return new ToolResult
            {
                Success = true,
                Data = $"已恢复到副本 {timestamp}。请调用 F_get_document_content 刷新快照与缓存。"
            };
        }

        public static string FormatListTextForCurrentDocument()
        {
            string docUuid = DocumentState.CurrDocUuid;
            if (string.IsNullOrEmpty(docUuid))
            {
                return "当前文档尚无操作记录。";
            }

            DocumentCheckpointSession session = GetSession(docUuid);
            return OperationLogFormatter.FormatListText(session);
        }

        public static void ClearOnDocumentClose(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            string docUuid = DocumentIdentity.TryResolveUuid(document);
            if (string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            ClearSessionForDocument(docUuid);
            System.Diagnostics.Debug.WriteLine($"[Checkpoint] clear on DocumentClose uuid={docUuid}");
        }

        public static void ClearSessionForDocument(string docUuid)
        {
            if (string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            Sessions.Remove(docUuid);

            string dir = GetDocCheckpointDirectory(docUuid);
            if (Directory.Exists(dir))
            {
                try
                {
                    Directory.Delete(dir, recursive: true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Checkpoint] delete dir failed uuid={docUuid}: {ex.Message}");
                }
            }
        }

        public static void ClearAllOnShutdown()
        {
            var uuids = Sessions.Keys.ToList();
            foreach (string uuid in uuids)
            {
                ClearSessionForDocument(uuid);
            }

            Sessions.Clear();
        }

        public static void ClearOrphanUuidFoldersOnShutdown()
        {
            if (_wordApplication == null)
            {
                return;
            }

            string root = GetCheckpointRootDirectory();
            if (!Directory.Exists(root))
            {
                return;
            }

            var liveUuids = new HashSet<string>(StringComparer.Ordinal);
            foreach (Word.Document doc in _wordApplication.Documents)
            {
                string uuid = DocumentIdentity.TryResolveUuid(doc);
                if (!string.IsNullOrEmpty(uuid))
                {
                    liveUuids.Add(uuid);
                }
            }

            foreach (string dir in Directory.GetDirectories(root))
            {
                string name = Path.GetFileName(dir);
                if (!liveUuids.Contains(name))
                {
                    try
                    {
                        Directory.Delete(dir, recursive: true);
                        System.Diagnostics.Debug.WriteLine($"[Checkpoint] clear orphan dir uuid={name}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[Checkpoint] orphan dir delete failed: {ex.Message}");
                    }
                }
            }
        }

        private static void AppendCheckpointEntry(DocumentCheckpointSession session, string timestamp)
        {
            session.Entries.Add(new OperationLogEntry
            {
                Kind = OperationLogEntryKind.Checkpoint,
                Timestamp = timestamp
            });
            session.CheckpointTimestampsInOrder.Add(timestamp);
        }

        private static void AppendProcessActionsLogs(
            DocumentCheckpointSession session,
            string toolName,
            IReadOnlyDictionary<string, object> parameters)
        {
            if (parameters == null || !parameters.TryGetValue("action", out object actionObj))
            {
                return;
            }

            if (!(actionObj is IEnumerable actionList))
            {
                return;
            }

            foreach (object item in actionList)
            {
                string json = OperationLogFormatter.FormatArgumentsJson(item);
                session.Entries.Add(new OperationLogEntry
                {
                    Kind = OperationLogEntryKind.Action,
                    ToolName = toolName,
                    ActionLineBody = json
                });
            }
        }

        private static void EvictOldestCheckpoint(DocumentCheckpointSession session, string docUuid)
        {
            if (session.CheckpointTimestampsInOrder.Count == 0)
            {
                return;
            }

            string oldest = session.CheckpointTimestampsInOrder[0];
            RemoveCheckpointAndFollowingActions(docUuid, oldest, deleteNewerFiles: false);
            DeleteCheckpointFile(docUuid, oldest);

            System.Diagnostics.Debug.WriteLine($"[Checkpoint] evict oldest timestamp={oldest}");
        }

        private static int TruncateLogAfterCheckpoint(string docUuid, string timestamp)
        {
            DocumentCheckpointSession session = GetSession(docUuid);
            if (session == null)
            {
                return 0;
            }

            int checkpointIndex = FindCheckpointEntryIndex(session, timestamp);
            if (checkpointIndex < 0)
            {
                return 0;
            }

            int tsIndex = session.CheckpointTimestampsInOrder.IndexOf(timestamp);
            if (tsIndex < 0)
            {
                return 0;
            }

            var toRemoveTimestamps = session.CheckpointTimestampsInOrder
                .Skip(tsIndex + 1)
                .ToList();

            foreach (string ts in toRemoveTimestamps)
            {
                DeleteCheckpointFile(docUuid, ts);
                session.CheckpointTimestampsInOrder.Remove(ts);
            }

            int removeFrom = checkpointIndex + 1;
            int removed = session.Entries.Count - removeFrom;
            if (removed > 0)
            {
                session.Entries.RemoveRange(removeFrom, removed);
            }

            return removed;
        }

        private static void RemoveCheckpointAndFollowingActions(
            string docUuid,
            string timestamp,
            bool deleteNewerFiles)
        {
            DocumentCheckpointSession session = GetSession(docUuid);
            if (session == null)
            {
                return;
            }

            int checkpointIndex = FindCheckpointEntryIndex(session, timestamp);
            if (checkpointIndex < 0)
            {
                session.CheckpointTimestampsInOrder.Remove(timestamp);
                DeleteCheckpointFile(docUuid, timestamp);
                return;
            }

            int nextCheckpointIndex = -1;
            for (int i = checkpointIndex + 1; i < session.Entries.Count; i++)
            {
                if (session.Entries[i].Kind == OperationLogEntryKind.Checkpoint)
                {
                    nextCheckpointIndex = i;
                    break;
                }
            }

            int removeCount = nextCheckpointIndex >= 0
                ? nextCheckpointIndex - checkpointIndex
                : session.Entries.Count - checkpointIndex;

            session.Entries.RemoveRange(checkpointIndex, removeCount);
            session.CheckpointTimestampsInOrder.Remove(timestamp);
            DeleteCheckpointFile(docUuid, timestamp);
        }

        private static int FindCheckpointEntryIndex(DocumentCheckpointSession session, string timestamp)
        {
            for (int i = 0; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && string.Equals(entry.Timestamp, timestamp, StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }

        private static void DeleteCheckpointFile(string docUuid, string timestamp)
        {
            string path = GetCheckpointFilePath(docUuid, timestamp);
            if (!File.Exists(path))
            {
                return;
            }

            try
            {
                File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] delete file failed: {ex.Message}");
            }
        }

        private static string GenerateUniqueTimestamp(string dir)
        {
            string baseTs = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            string timestamp = baseTs;
            int suffix = 1;
            while (File.Exists(Path.Combine(dir, $"checkpoint_{timestamp}.docx")))
            {
                timestamp = $"{baseTs}_{suffix}";
                suffix++;
            }

            return timestamp;
        }

        private static bool HasAnyFailedActionResult(object resultsObj)
        {
            if (resultsObj is IEnumerable list && !(resultsObj is string))
            {
                foreach (object item in list)
                {
                    if (TryGetSuccessFlag(item, out bool success) && !success)
                    {
                        return true;
                    }
                }

                return false;
            }

            if (resultsObj is Dictionary<string, object> dict)
            {
                foreach (object value in dict.Values)
                {
                    if (HasAnyFailedActionResult(value))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryGetSuccessFlag(object item, out bool success)
        {
            success = true;
            if (item is Dictionary<string, object> dict && dict.TryGetValue("success", out object flag))
            {
                success = Convert.ToBoolean(flag);
                return true;
            }

            if (item is IDictionary legacy && legacy.Contains("success"))
            {
                success = Convert.ToBoolean(legacy["success"]);
                return true;
            }

            return false;
        }

        private static CheckpointBeforeResult Fail(string error)
        {
            return new CheckpointBeforeResult { Success = false, Error = error };
        }
    }

    /// <summary>
    /// 将活动文档另存为副本文件。Word Interop 的 Document.SaveCopyAs 在 Word 中未实现（会抛 COM 0x800A1704），
    /// 使用 IPersistFile.Save 或 SaveAs2+路径恢复 作为替代。
    /// </summary>
    public static class DocumentCheckpointSaveHelper
    {
        public static void SaveDocumentCopy(Word.Document document, string destinationPath)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            if (string.IsNullOrWhiteSpace(destinationPath))
            {
                throw new ArgumentException("destinationPath is required", nameof(destinationPath));
            }

            string path = Path.GetFullPath(destinationPath);
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (TrySaveViaPersistFile(document, path))
            {
                return;
            }

            if (TrySaveViaTemporarySaveAs2(document, path))
            {
                return;
            }

            throw new InvalidOperationException(
                "无法创建文档副本（IPersistFile 与 SaveAs2 均失败）。请先保存文档后重试。");
        }

        private static bool TrySaveViaPersistFile(Word.Document document, string path)
        {
            try
            {
                IPersistFile persistFile = GetPersistFile(document);
                if (persistFile == null)
                {
                    return false;
                }

                persistFile.Save(path, false);

                if (!File.Exists(path) || new FileInfo(path).Length == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[Checkpoint] IPersistFile.Save produced empty/missing file");
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"[Checkpoint] IPersistFile.Save ok path={path}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[Checkpoint] IPersistFile.Save failed: {ex.Message} (HRESULT={ExtractHResult(ex)})");
                return false;
            }
        }

        private static bool TrySaveViaTemporarySaveAs2(Word.Document document, string path)
        {
            string originalFullName = null;
            try
            {
                originalFullName = document.FullName;
            }
            catch
            {
                originalFullName = "";
            }

            bool hadOriginalPath = !string.IsNullOrEmpty(originalFullName)
                && !originalFullName.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase)
                && File.Exists(originalFullName);

            if (!hadOriginalPath)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[Checkpoint] SaveAs2 fallback skipped: document has no saved path (use IPersistFile)");
                return false;
            }

            try
            {
                document.SaveAs2(path, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);

                if (!File.Exists(path) || new FileInfo(path).Length == 0)
                {
                    return false;
                }

                document.SaveAs2(originalFullName, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);

                System.Diagnostics.Debug.WriteLine(
                    $"[Checkpoint] SaveAs2 copy ok path={path} restored={hadOriginalPath}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] SaveAs2 copy failed: {ex.Message}");

                try
                {
                    document.SaveAs2(originalFullName, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
                }
                catch (Exception restoreEx)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[Checkpoint] SaveAs2 restore original failed: {restoreEx.Message}");
                }

                return false;
            }
        }

        private static IPersistFile GetPersistFile(Word.Document document)
        {
            if (document == null || !Marshal.IsComObject(document))
            {
                return null;
            }

            try
            {
                return (IPersistFile)document;
            }
            catch (InvalidCastException)
            {
                IntPtr unk = IntPtr.Zero;
                try
                {
                    unk = Marshal.GetIUnknownForObject(document);
                    return (IPersistFile)Marshal.GetObjectForIUnknown(unk);
                }
                catch
                {
                    return null;
                }
                finally
                {
                    if (unk != IntPtr.Zero)
                    {
                        Marshal.Release(unk);
                    }
                }
            }
        }

        private static string ExtractHResult(Exception ex)
        {
            if (ex is COMException com)
            {
                return "0x" + com.ErrorCode.ToString("X8");
            }

            return "";
        }
    }
}
