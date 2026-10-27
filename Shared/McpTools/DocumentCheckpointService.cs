using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using WordAddIn1.DocumentHost;
using WordAddIn1.OpenFiles;
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

        public static string GetCheckpointFilePath(string docUuid, string timestamp, string extension = "docx")
        {
            string ext = string.IsNullOrEmpty(extension) ? "docx" : extension.TrimStart('.');
            return Path.Combine(GetDocCheckpointDirectory(docUuid), $"checkpoint_{timestamp}.{ext}");
        }

        public static string FindCheckpointFile(string docUuid, string timestamp)
        {
            foreach (string ext in new[] { "docx", "xlsx", "pptx" })
            {
                string path = GetCheckpointFilePath(docUuid, timestamp, ext);
                if (File.Exists(path))
                {
                    return path;
                }
            }

            return GetCheckpointFilePath(docUuid, timestamp, "docx");
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
            Dictionary<string, object> args = ToArgsDictionary(parameters);
            string explicitId = ChannelContext.TryGetChannelIdFromParameters(args);
            bool hasExplicitChannel = !string.IsNullOrEmpty(explicitId);
            HostProbe.Log("ckpt-enter", "tool=" + toolName + " " + HostProbe.ArgsBrief(args)
                + " hasExplicit=" + hasExplicitChannel + " wordApp=" + (_wordApplication != null));

            if (!TryResolveLiveTarget(args, hasExplicitChannel, out CheckpointLiveTarget live, out string resolveHint))
            {
                return Fail(
                    string.IsNullOrEmpty(resolveHint)
                        ? "无可用文档渠道；请确认已打开文字稿、工作簿或演示文稿，或传入 channel_id"
                        : resolveHint,
                    "ckpt-no-doc");
            }

            CheckpointCaret caret = live.Adapter.CaptureCaret(live);

            try
            {
                live.Adapter.PrepareForSave(live);
            }
            catch (Exception ex)
            {
                HostProbe.Log("ckpt-bind-throw", HostProbe.Ex(ex));
                return Fail(ex.Message, "ckpt-bind-throw");
            }

            if (string.IsNullOrEmpty(live.DocUuid) && live.Live is Word.Document wordDoc)
            {
                live.DocUuid = DocumentState.EnsureCurrDocUuid(wordDoc);
            }

            if (string.IsNullOrEmpty(live.DocUuid))
            {
                return Fail("当前文档尚未绑定 doc_uuid", "ckpt-no-uuid");
            }

            string dir = GetDocCheckpointDirectory(live.DocUuid);
            Directory.CreateDirectory(dir);

            string timestamp = GenerateUniqueTimestamp(dir);
            string path = GetCheckpointFilePath(live.DocUuid, timestamp, live.Adapter.CopyExtension);

            try
            {
                live.Adapter.SaveCopy(live, path);
            }
            catch (Exception ex)
            {
                HostProbe.Log("ckpt-save-fail", HostProbe.Ex(ex));
                return Fail($"创建文档副本失败: {ex.Message}", "ckpt-save");
            }

            DocumentCheckpointSession session = GetOrCreateSession(live.DocUuid);
            AppendCheckpointEntry(session, timestamp);
            session.CaretByTimestamp[timestamp] = caret;

            if (session.CheckpointTimestampsInOrder.Count > MaxCheckpointSlots)
            {
                EvictOldestCheckpoint(session, live.DocUuid);
            }

            HostProbe.Log("ckpt-ok", "path=" + path + " ts=" + timestamp);
            System.Diagnostics.Debug.WriteLine(
                $"[Checkpoint] copy saved path={path} timestamp={timestamp} historyCount={session.CheckpointTimestampsInOrder.Count}");

            return new CheckpointBeforeResult
            {
                Success = true,
                Timestamp = timestamp,
                DocUuid = live.DocUuid,
                ChannelId = live.ChannelId
            };
        }

        private static bool TryResolveLiveTarget(
            Dictionary<string, object> args,
            bool hasExplicitChannel,
            out CheckpointLiveTarget live,
            out string resolveHint)
        {
            live = null;
            resolveHint = null;

            if (ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string channelError)
                && channel != null)
            {
                if (CheckpointKindAdapters.TryFromChannel(channel, out live, out resolveHint))
                {
                    return true;
                }

                if (hasExplicitChannel)
                {
                    return false;
                }
            }
            else if (!string.IsNullOrEmpty(channelError))
            {
                resolveHint = channelError;
                if (hasExplicitChannel)
                {
                    return false;
                }
            }

            if (!TryResolveDocumentForCheckpoint(args, out Word.Document activeDoc, out resolveHint)
                && LooksLikeNoOpenDocument(resolveHint))
            {
                TryActivateChannelDocument(args);
                TryResolveDocumentForCheckpoint(args, out activeDoc, out resolveHint);
            }

            if (activeDoc == null && !hasExplicitChannel && _wordApplication != null)
            {
                try
                {
                    activeDoc = _wordApplication.ActiveDocument;
                }
                catch (Exception ex)
                {
                    resolveHint = string.IsNullOrEmpty(resolveHint)
                        ? "无法读取活动文档: " + ex.Message
                        : resolveHint;
                    return false;
                }
            }

            if (activeDoc == null)
            {
                return false;
            }

            string uuid = null;
            string channelId = ChannelContext.TryGetChannelIdFromParameters(args);
            IOperationChannel wordChannel = null;
            if (ChannelContext.TryResolveChannel(args, out wordChannel, out _)
                && wordChannel != null)
            {
                channelId = wordChannel.ChannelId;
                uuid = ChannelRegistry.TryGetChannelDocUuid(wordChannel);
            }

            live = new CheckpointLiveTarget
            {
                Adapter = CheckpointKindAdapters.Word,
                Channel = wordChannel,
                ChannelId = channelId,
                DocUuid = uuid,
                Live = activeDoc
            };
            return true;
        }

        private static bool TryResolveDocumentForCheckpoint(
            Dictionary<string, object> args,
            out Word.Document document,
            out string resolveHint)
        {
            document = null;
            resolveHint = null;
            try
            {
                if (DocumentHostAdapter.TryResolveInteropDocument(
                        args,
                        _wordApplication,
                        out InteropDocumentHandle handle,
                        out ToolResult resolveError))
                {
                    document = handle.Document;
                    TryCacheApplicationFromDocument(document);
                    HostProbe.Log("ckpt-resolve-ok", "host=" + handle.HostName + " ch=" + handle.ChannelId);
                    HostProbe.Touch("ckpt-resolve-doc", document);
                    return document != null;
                }

                resolveHint = resolveError?.Error;
                HostProbe.Log("ckpt-resolve-false", "hint=" + resolveHint);
                return false;
            }
            catch (Exception ex)
            {
                resolveHint = ex.Message;
                HostProbe.Log("ckpt-resolve-throw", HostProbe.Ex(ex));
                return false;
            }
        }

        private static void TryActivateChannelDocument(Dictionary<string, object> args)
        {
            if (!HostCallbacks.ShouldStealForeground())
            {
                return;
            }

            try
            {
                if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out _)
                    || channel == null)
                {
                    return;
                }

                if (channel is WordChannel word && word.TryGetLiveDocument(out Word.Document wordDoc))
                {
                    TryActivateInteropDocument(wordDoc);
                    return;
                }

                if (channel is WpsChannel wps && wps.TryGetLiveDocument(out object wpsDoc))
                {
                    if (DocumentHostCompat.TryAsWordDocument(wpsDoc, out Word.Document compat))
                    {
                        TryActivateInteropDocument(compat);
                    }
                    else
                    {
                        DocumentHostCompat.TryActivateLateBound(wpsDoc);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] activate channel skipped: {ex.Message}");
            }
        }

        private static void TryActivateInteropDocument(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            try
            {
                document.Activate();
            }
            catch (Exception)
            {
                DocumentHostCompat.TryActivateLateBound(document);
            }
        }

        private static bool LooksLikeNoOpenDocument(string text)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf("没有打开的文档", StringComparison.Ordinal) >= 0;
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
            if (string.Equals(toolName, "F_process_comments", StringComparison.Ordinal))
            {
                return IsCheckpointCleanupOnly(result);
            }

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

        public static void RollbackFailedMutation(string docUuid, string timestamp, string channelId = null)
        {
            if (string.IsNullOrEmpty(docUuid) || string.IsNullOrEmpty(timestamp))
            {
                return;
            }

            try
            {
                string path = FindCheckpointFile(docUuid, timestamp);
                if (File.Exists(path)
                    && TryResolveLiveTargetForRestore(docUuid, channelId, out CheckpointLiveTarget live))
                {
                    HostProbe.Log(
                        "rollback-restore",
                        "uuid=" + docUuid + " ch=" + (channelId ?? "") + " path=" + path);
                    live.Adapter.RestoreCopy(live, path);
                    TryRestoreStoredCaret(docUuid, timestamp, live);
                }
                else
                {
                    HostProbe.Log(
                        "rollback-skip-restore",
                        "uuid=" + docUuid + " ch=" + (channelId ?? "")
                        + " file=" + File.Exists(path));
                }
            }
            catch (Exception ex)
            {
                HostProbe.Log("rollback-warn", HostProbe.Ex(ex));
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] rollback restore warning: {ex.Message}");
            }

            try
            {
                DocumentState.ClearAfterCheckpointRestore();
                RemoveCheckpointAndFollowingActions(docUuid, timestamp, deleteNewerFiles: false);
                DeleteCheckpointFile(docUuid, timestamp);
            }
            catch (Exception ex)
            {
                HostProbe.Log("rollback-cleanup-warn", HostProbe.Ex(ex));
            }

            System.Diagnostics.Debug.WriteLine($"[Checkpoint] rollback failed invoke timestamp={timestamp}");
        }

        private static bool TryResolveLiveTargetForRestore(
            string docUuid,
            string channelId,
            out CheckpointLiveTarget live)
        {
            live = null;
            IOperationChannel channel = null;
            if (!string.IsNullOrEmpty(channelId))
            {
                ChannelRegistry.TryGet(channelId, out channel);
            }

            if (channel == null)
            {
                ChannelRegistry.TryGetChannelByDocUuid(docUuid, out channel);
            }

            if (channel != null && CheckpointKindAdapters.TryFromChannel(channel, out live, out _))
            {
                if (string.IsNullOrEmpty(live.DocUuid))
                {
                    live.DocUuid = docUuid;
                }

                return true;
            }

            Word.Document wordDoc = TryResolveLiveDocumentForRollback(docUuid, channelId);
            if (wordDoc == null)
            {
                return false;
            }

            live = new CheckpointLiveTarget
            {
                Adapter = CheckpointKindAdapters.Word,
                ChannelId = channelId,
                DocUuid = docUuid,
                Live = wordDoc
            };
            return true;
        }

        private static void TryRestoreStoredCaret(string docUuid, string timestamp, CheckpointLiveTarget live)
        {
            DocumentCheckpointSession session = GetSession(docUuid);
            if (session == null
                || !session.CaretByTimestamp.TryGetValue(timestamp, out CheckpointCaret caret)
                || caret == null)
            {
                return;
            }

            try
            {
                live.Adapter.RestoreCaret(live, caret);
            }
            catch (Exception ex)
            {
                HostProbe.Log("restore-caret-warn", HostProbe.Ex(ex));
            }
        }

        private static Word.Document TryResolveLiveDocumentForRollback(string docUuid, string channelId)
        {
            if (!string.IsNullOrEmpty(channelId)
                && ChannelRegistry.TryGet(channelId, out IOperationChannel byId)
                && byId != null)
            {
                Word.Document fromChannel = TryLiveDocumentFromChannel(byId);
                if (fromChannel != null)
                {
                    return fromChannel;
                }
            }

            return TryResolveLiveDocumentByUuid(docUuid);
        }

        private static Word.Document TryResolveLiveDocumentByUuid(string docUuid)
        {
            if (!ChannelRegistry.TryGetChannelByDocUuid(docUuid, out IOperationChannel channel) || channel == null)
            {
                return null;
            }

            return TryLiveDocumentFromChannel(channel);
        }

        private static Word.Document TryLiveDocumentFromChannel(IOperationChannel channel)
        {
            if (channel is WordChannel word && word.TryGetLiveDocument(out Word.Document wordDoc))
            {
                return wordDoc;
            }

            if (channel is WpsChannel wps && wps.TryGetLiveDocument(out object wpsDoc)
                && DocumentHostCompat.TryAsWordDocument(wpsDoc, out Word.Document compat))
            {
                return compat;
            }

            return null;
        }

        /// <summary>
        /// 回滚 Open 必须走本稿宿主的 Application。WPS 兼容 RCW 的 .Application 双开时会落到旁边 Word。
        /// </summary>
        private static Word.Application TryApplicationForRestore(Word.Document target, string channelId)
        {
            if (!string.IsNullOrEmpty(channelId)
                && ChannelRegistry.TryGet(channelId, out IOperationChannel channel)
                && channel is WpsChannel wps
                && wps.TryGetLiveDocument(out object wpsDoc))
            {
                Word.Application late = TryAsWordApplication(WpsCom.GetProperty(wpsDoc, "Application"));
                if (late != null)
                {
                    return late;
                }
            }

            try
            {
                return target?.Application;
            }
            catch (Exception ex)
            {
                HostProbe.Log("rollback-app", HostProbe.Ex(ex));
                return null;
            }
        }

        private static Word.Application TryAsWordApplication(object app)
        {
            if (app == null)
            {
                return null;
            }

            Word.Application typed = app as Word.Application;
            if (typed != null)
            {
                return typed;
            }

            try
            {
                return (Word.Application)app;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static ToolResult RestoreCheckpoint(string timestamp, Word.Document document = null)
        {
            CheckpointLiveTarget live = null;
            if (document != null)
            {
                live = new CheckpointLiveTarget
                {
                    Adapter = CheckpointKindAdapters.Word,
                    Live = document
                };
                try
                {
                    DocumentState.BindAndActivate(document);
                    live.DocUuid = DocumentState.CurrDocUuid;
                }
                catch (Exception)
                {
                }
            }

            return RestoreCheckpoint(timestamp, live);
        }

        public static ToolResult RestoreCheckpoint(string timestamp, CheckpointLiveTarget live)
        {
            if (live == null || live.Live == null || live.Adapter == null)
            {
                return new ToolResult { Success = false, Error = "没有活动的文档窗口" };
            }

            if (string.IsNullOrEmpty(live.DocUuid) && live.Channel != null)
            {
                live.DocUuid = ChannelRegistry.TryGetChannelDocUuid(live.Channel);
            }

            if (string.IsNullOrEmpty(live.DocUuid) && live.Live is Word.Document wordDoc)
            {
                DocumentState.BindAndActivate(wordDoc);
                live.DocUuid = DocumentState.CurrDocUuid;
            }

            if (string.IsNullOrEmpty(live.DocUuid))
            {
                return new ToolResult { Success = false, Error = "当前文档尚未绑定 doc_uuid" };
            }

            DocumentCheckpointSession session = GetSession(live.DocUuid);
            if (session == null || !session.CheckpointTimestampsInOrder.Contains(timestamp))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = $"副本 timestamp '{timestamp}' 不存在或已被淘汰。请先调用 F_list_document_operations 查看可用副本。"
                };
            }

            string path = FindCheckpointFile(live.DocUuid, timestamp);
            if (!File.Exists(path))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = $"副本文件不存在: checkpoint_{timestamp}。请先调用 F_list_document_operations。"
                };
            }

            try
            {
                live.Adapter.RestoreCopy(live, path);
                TryRestoreStoredCaret(live.DocUuid, timestamp, live);
            }
            catch (Exception ex)
            {
                Exception inner = ex;
                while (inner.InnerException != null)
                {
                    inner = inner.InnerException;
                }

                return new ToolResult
                {
                    Success = false,
                    Error = "恢复副本失败: " + (string.IsNullOrEmpty(inner.Message) ? ex.Message : inner.Message)
                };
            }

            DocumentState.ClearAfterCheckpointRestore();
            int removed = TruncateLogAfterCheckpoint(live.DocUuid, timestamp);

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
            return FormatListTextForDocUuid(DocumentState.CurrDocUuid);
        }

        public static string FormatListTextForDocUuid(string docUuid)
        {
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
                session.CaretByTimestamp.Remove(ts);
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
            foreach (string ext in new[] { "docx", "xlsx", "pptx" })
            {
                string path = GetCheckpointFilePath(docUuid, timestamp, ext);
                if (!File.Exists(path))
                {
                    continue;
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

            DocumentCheckpointSession session = GetSession(docUuid);
            session?.CaretByTimestamp.Remove(timestamp);
        }

        private static string GenerateUniqueTimestamp(string dir)
        {
            string baseTs = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            string timestamp = baseTs;
            int suffix = 1;
            while (CheckpointFileExists(dir, timestamp))
            {
                timestamp = $"{baseTs}_{suffix}";
                suffix++;
            }

            return timestamp;
        }

        private static bool CheckpointFileExists(string dir, string timestamp)
        {
            foreach (string ext in new[] { "docx", "xlsx", "pptx" })
            {
                if (File.Exists(Path.Combine(dir, $"checkpoint_{timestamp}.{ext}")))
                {
                    return true;
                }
            }

            return false;
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

        private static CheckpointBeforeResult Fail(string error, string site = null)
        {
            string text = error ?? "";
            if (!string.IsNullOrEmpty(site))
            {
                HostProbe.Log("ckpt-fail", "site=" + site + " err=" + text);
                text = "[" + site + "] " + text;
            }

            return new CheckpointBeforeResult { Success = false, Error = text };
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

            if (TrySaveViaFileCopy(document, path))
            {
                return;
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
                "无法创建文档副本（文件拷贝、IPersistFile 与 SaveAs2 均失败）。请先保存文档后重试。");
        }

        /// <summary>
        /// WPS 的 IPersistFile.Save 常把「另存副本」做成改路径/拆掉 ActiveDocument，
        /// 下一轮改字就会报「没有打开的文档」。已落盘的稿先 Save 再拷文件。
        /// </summary>
        private static bool TrySaveViaFileCopy(Word.Document document, string path)
        {
            string originalFullName;
            try
            {
                originalFullName = document.FullName;
            }
            catch (Exception)
            {
                return false;
            }

            if (string.IsNullOrEmpty(originalFullName)
                || originalFullName.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase)
                || !File.Exists(originalFullName))
            {
                return false;
            }

            try
            {
                try
                {
                    document.Save();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[Checkpoint] in-place Save before copy skipped: " + ex.Message);
                }

                File.Copy(originalFullName, path, overwrite: true);
                if (!File.Exists(path) || new FileInfo(path).Length == 0)
                {
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"[Checkpoint] File.Copy ok path={path}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] File.Copy failed: {ex.Message}");
                return false;
            }
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
