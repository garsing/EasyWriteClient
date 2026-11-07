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
        private const int MaxOrdinaryPerBranch = 10;
        private const int MaxUserPerBranch = 20;
        public const string RoundMutatingMessage = "正在修改，请先停止或等这一轮结束";

        private static readonly Dictionary<string, DocumentCheckpointSession> Sessions =
            new Dictionary<string, DocumentCheckpointSession>(StringComparer.Ordinal);

        private static readonly Dictionary<string, string> RoundMutatingChannelByDocId =
            new Dictionary<string, string>(StringComparer.Ordinal);

        private static readonly Dictionary<string, string> RoundBeforeUserTimestampByDocId =
            new Dictionary<string, string>(StringComparer.Ordinal);

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
                session = CheckpointHistoryStore.TryLoad(docUuid);
                if (session == null)
                {
                    session = NewSession(docUuid);
                }

                Sessions[docUuid] = session;
            }

            return session;
        }

        public static DocumentCheckpointSession PeekSession(string docId)
        {
            if (string.IsNullOrEmpty(docId))
            {
                return null;
            }

            if (Sessions.TryGetValue(docId, out DocumentCheckpointSession session))
            {
                return session;
            }

            session = CheckpointHistoryStore.TryLoad(docId);
            if (session != null)
            {
                Sessions[docId] = session;
            }

            return session;
        }

        public static void StartFreshIdentity(string docId, string fullName, object live)
        {
            if (string.IsNullOrEmpty(docId))
            {
                return;
            }

            var session = NewSession(docId);
            ApplyLiveMetadata(session, live, fullName);
            Sessions[docId] = session;
            Persist(session);
        }

        private static DocumentCheckpointSession NewSession(string docId)
        {
            var session = new DocumentCheckpointSession
            {
                DocUuid = docId,
                CurrentBranch = 1
            };
            session.Branches.Add(new CheckpointBranch { Id = 1 });
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

            string docId = BindCheckpointIdentity(live, create: true, out string identityError);
            if (string.IsNullOrEmpty(docId))
            {
                return Fail(
                    string.IsNullOrEmpty(identityError)
                        ? "创建文档副本失败: 无法写入文档身份 EasyWriteDocId"
                        : identityError,
                    "ckpt-no-uuid");
            }

            live.DocUuid = docId;
            DocumentCheckpointSession session = GetOrCreateSession(docId);
            ApplyLiveMetadata(session, live.Live, EasyWriteDocIdStore.TryReadFullName(live.Live), live.Adapter.CopyExtension);
            MaybeFork(session);

            string dir = GetDocCheckpointDirectory(docId);
            Directory.CreateDirectory(dir);

            string timestamp = GenerateUniqueTimestamp(dir);
            string path = GetCheckpointFilePath(docId, timestamp, live.Adapter.CopyExtension);

            try
            {
                live.Adapter.SaveCopy(live, path);
            }
            catch (Exception ex)
            {
                HostProbe.Log("ckpt-save-fail", HostProbe.Ex(ex));
                RevertBranchIfEmpty(session);
                Persist(session);
                return Fail($"创建文档副本失败: {ex.Message}", "ckpt-save");
            }

            string hash = CheckpointContentHash.FromFile(path);
            if (!RoundBeforeUserTimestampByDocId.ContainsKey(docId))
            {
                AppendCheckpointEntry(session, timestamp, session.CurrentBranch, userRollback: true, operations: null, hash);
                RoundBeforeUserTimestampByDocId[docId] = timestamp;
            }

            AppendCheckpointEntry(session, timestamp, session.CurrentBranch, userRollback: false, operations: null, hash);
            session.CaretByTimestamp[timestamp] = caret;
            EvictOldestOrdinary(session, docId);
            RoundMutatingChannelByDocId[docId] = live.ChannelId ?? "";
            Persist(session);

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
            IReadOnlyDictionary<string, object> parameters,
            string checkpointDocId)
        {
            if (string.IsNullOrEmpty(checkpointDocId))
            {
                return;
            }

            DocumentCheckpointSession session = GetOrCreateSession(checkpointDocId);
            int branch = session.CurrentBranch > 0 ? session.CurrentBranch : 1;

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
                    Branch = branch,
                    ToolName = toolName,
                    ActionLineBody = json
                });
            }

            Persist(session);
            System.Diagnostics.Debug.WriteLine($"[Checkpoint] append action tool={toolName}");
        }

        public static bool IsInvokeMutationFailed(ToolResult result, string toolName)
        {
            if (string.Equals(toolName, "F_process_comments", StringComparison.Ordinal))
            {
                return IsCheckpointCleanupOnly(result);
            }

            if (IsDocumentActionsOutcome(toolName))
            {
                if (IsCheckpointCleanupOnly(result))
                {
                    return true;
                }

                int? executed = GetIntProperty(result?.Data, "executed_actions");
                if (executed.HasValue && executed.Value > 0)
                {
                    return false;
                }

                return result == null || !result.Success;
            }

            if (result == null || !result.Success)
            {
                return true;
            }

            return false;
        }

        private static bool IsDocumentActionsOutcome(string toolName)
        {
            return string.Equals(toolName, "F_process_document_actions", StringComparison.Ordinal)
                || string.Equals(toolName, "F_insert_document_from_file", StringComparison.Ordinal)
                || string.Equals(toolName, "F_replace_document_from_file", StringComparison.Ordinal);
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

            RemoveFailedInvokeCheckpoint(docUuid, timestamp);
            DocumentCheckpointSession session = GetSession(docUuid);
            RevertBranchIfEmpty(session);
            Persist(session);
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
                RemoveFailedInvokeCheckpoint(docUuid, timestamp);
                DocumentCheckpointSession session = GetSession(docUuid);
                RevertBranchIfEmpty(session);
                Persist(session);
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

            DocumentCheckpointSession session = GetOrCreateSession(live.DocUuid);
            if (session == null || !session.CheckpointTimestampsInOrder.Contains(timestamp))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = $"副本 timestamp '{timestamp}' 不存在或已被淘汰。请先调用 F_list_document_operations 查看可用副本。"
                };
            }

            OperationLogEntry chosen = FindPreferredCheckpoint(session, timestamp);
            if (chosen != null
                && !chosen.UserRollback
                && (chosen.Retired || chosen.Branch != LatestBranchId(session)))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = "这份副本已经不能恢复"
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
            OperationLogEntry restored = FindCheckpointEntry(session, timestamp);
            if (restored != null)
            {
                session.PositionBranch = restored.Branch > 0 ? restored.Branch : session.CurrentBranch;
                session.PositionTimestamp = timestamp;
                Persist(session);
            }

            System.Diagnostics.Debug.WriteLine(
                $"[Checkpoint] restore timestamp={timestamp} positionBranch={session.PositionBranch}");

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

            DocumentCheckpointSession session = GetOrCreateSession(docUuid);
            return OperationLogFormatter.FormatListText(session);
        }

        public static void ClearOnDocumentClose(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            string fileId = EasyWriteDocIdStore.TryRead(document);
            if (!string.IsNullOrEmpty(fileId))
            {
                Sessions.Remove(fileId);
                RoundMutatingChannelByDocId.Remove(fileId);
            }

            string docUuid = DocumentIdentity.TryResolveUuid(document);
            if (!string.IsNullOrEmpty(docUuid))
            {
                Sessions.Remove(docUuid);
            }

            System.Diagnostics.Debug.WriteLine($"[Checkpoint] unload on DocumentClose fileId={fileId}");
        }

        public static void ClearSessionForDocument(string docUuid)
        {
            if (string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            Sessions.Remove(docUuid);
            RoundMutatingChannelByDocId.Remove(docUuid);
        }

        public static void ClearAllOnShutdown()
        {
            try
            {
                MarkOrdinaryRetiredOnShutdown();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Checkpoint] retire on shutdown failed: " + ex.Message);
            }

            Sessions.Clear();
            RoundMutatingChannelByDocId.Clear();
            RoundBeforeUserTimestampByDocId.Clear();
        }

        public static void MarkOrdinaryRetiredOnShutdown()
        {
            var roots = new List<string>();
            try
            {
                roots.Add(GetCheckpointRootDirectory());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Checkpoint] checkpoint root failed: " + ex.Message);
            }

            string tempRoot = Path.Combine(Path.GetTempPath(), "EasyWrite", "document_checkpoints");
            if (Directory.Exists(tempRoot))
            {
                bool seen = false;
                foreach (string root in roots)
                {
                    if (string.Equals(root, tempRoot, StringComparison.OrdinalIgnoreCase))
                    {
                        seen = true;
                        break;
                    }
                }

                if (!seen)
                {
                    roots.Add(tempRoot);
                }
            }

            foreach (string root in roots)
            {
                if (string.IsNullOrEmpty(root) || !Directory.Exists(root))
                {
                    continue;
                }

                foreach (string dir in Directory.GetDirectories(root))
                {
                    CheckpointHistoryStore.MarkOrdinaryRetired(Path.Combine(dir, "history.json"));
                }
            }
        }

        public static void ClearOrphanUuidFoldersOnShutdown()
        {
        }

        private static void AppendCheckpointEntry(
            DocumentCheckpointSession session,
            string timestamp,
            int branch,
            bool userRollback,
            List<CheckpointOperationRecord> operations,
            string hash)
        {
            var entry = new OperationLogEntry
            {
                Kind = OperationLogEntryKind.Checkpoint,
                Timestamp = timestamp,
                Branch = branch > 0 ? branch : 1,
                UserRollback = userRollback,
                Hash = hash ?? "",
                Retired = false
            };
            if (operations != null)
            {
                entry.Operations.AddRange(operations);
            }

            session.Entries.Add(entry);
            if (!session.CheckpointTimestampsInOrder.Contains(timestamp))
            {
                session.CheckpointTimestampsInOrder.Add(timestamp);
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
                    Branch = session.CurrentBranch > 0 ? session.CurrentBranch : 1,
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

        private static void RemoveFailedInvokeCheckpoint(string docUuid, string timestamp)
        {
            DocumentCheckpointSession session = GetSession(docUuid);
            if (session == null)
            {
                DeleteCheckpointFile(docUuid, timestamp);
                return;
            }

            int ordinary = FindCheckpointEntryIndex(session, timestamp, userRollback: false);
            if (ordinary >= 0)
            {
                RemoveOrdinaryEntry(session, docUuid, ordinary);
            }

            for (int i = session.Entries.Count - 1; i >= 0; i--)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && entry.UserRollback
                    && string.Equals(entry.Timestamp, timestamp, StringComparison.Ordinal))
                {
                    session.Entries.RemoveAt(i);
                }
            }

            session.CheckpointTimestampsInOrder.Remove(timestamp);
            DeleteCheckpointFile(docUuid, timestamp);
            if (RoundBeforeUserTimestampByDocId.TryGetValue(docUuid, out string paired)
                && string.Equals(paired, timestamp, StringComparison.Ordinal))
            {
                RoundBeforeUserTimestampByDocId.Remove(docUuid);
            }
        }

        private static void RemoveOrdinaryEntry(DocumentCheckpointSession session, string docId, int index)
        {
            if (session == null || index < 0 || index >= session.Entries.Count)
            {
                return;
            }

            string timestamp = session.Entries[index].Timestamp;
            int end = index + 1;
            while (end < session.Entries.Count
                && session.Entries[end].Kind != OperationLogEntryKind.Checkpoint)
            {
                end++;
            }

            session.Entries.RemoveRange(index, end - index);
            if (!CheckpointTimestampStillUsed(session, timestamp))
            {
                session.CheckpointTimestampsInOrder.Remove(timestamp);
                DeleteCheckpointFile(docId, timestamp);
            }
        }

        private static void RemoveSpecificCheckpoint(DocumentCheckpointSession session, string docId, int index)
        {
            if (session == null || index < 0 || index >= session.Entries.Count)
            {
                return;
            }

            string timestamp = session.Entries[index].Timestamp;
            session.Entries.RemoveAt(index);
            if (!CheckpointTimestampStillUsed(session, timestamp))
            {
                session.CheckpointTimestampsInOrder.Remove(timestamp);
                DeleteCheckpointFile(docId, timestamp);
            }
        }

        private static bool CheckpointTimestampStillUsed(DocumentCheckpointSession session, string timestamp)
        {
            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && string.Equals(entry.Timestamp, timestamp, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static int FindCheckpointEntryIndex(DocumentCheckpointSession session, string timestamp, bool userRollback)
        {
            for (int i = 0; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && entry.UserRollback == userRollback
                    && string.Equals(entry.Timestamp, timestamp, StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }

        private static OperationLogEntry FindPreferredCheckpoint(DocumentCheckpointSession session, string timestamp)
        {
            OperationLogEntry ordinary = null;
            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry.Kind != OperationLogEntryKind.Checkpoint
                    || !string.Equals(entry.Timestamp, timestamp, StringComparison.Ordinal))
                {
                    continue;
                }

                if (entry.UserRollback)
                {
                    return entry;
                }

                ordinary = entry;
            }

            return ordinary;
        }

        private static int LatestBranchId(DocumentCheckpointSession session)
        {
            int latest = 1;
            if (session == null)
            {
                return latest;
            }

            foreach (CheckpointBranch branch in session.Branches)
            {
                if (branch != null && branch.Id > latest)
                {
                    latest = branch.Id;
                }
            }

            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry != null && entry.Branch > latest)
                {
                    latest = entry.Branch;
                }
            }

            return latest;
        }

        private static int CountUsableOrdinary(DocumentCheckpointSession session, int branch)
        {
            int count = 0;
            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && !entry.UserRollback
                    && !entry.Retired
                    && entry.Branch == branch)
                {
                    count++;
                }
            }

            return count;
        }

        private static int FindOldestUsableOrdinary(DocumentCheckpointSession session, int branch)
        {
            for (int i = 0; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && !entry.UserRollback
                    && !entry.Retired
                    && entry.Branch == branch)
                {
                    return i;
                }
            }

            return -1;
        }

        private static int CountBillableUsers(DocumentCheckpointSession session, int branch)
        {
            int count = 0;
            bool seenFirst = false;
            string previousHash = null;
            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry.Kind != OperationLogEntryKind.Checkpoint
                    || !entry.UserRollback
                    || entry.Branch != branch)
                {
                    continue;
                }

                if (!seenFirst)
                {
                    seenFirst = true;
                    previousHash = entry.Hash ?? "";
                    continue;
                }

                if (!string.IsNullOrEmpty(entry.Hash)
                    && string.Equals(entry.Hash, previousHash, StringComparison.Ordinal))
                {
                    continue;
                }

                count++;
                previousHash = entry.Hash ?? "";
            }

            return count;
        }

        private static int FindOldestBillableUser(DocumentCheckpointSession session, int branch)
        {
            bool seenFirst = false;
            string previousHash = null;
            for (int i = 0; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind != OperationLogEntryKind.Checkpoint
                    || !entry.UserRollback
                    || entry.Branch != branch)
                {
                    continue;
                }

                if (!seenFirst)
                {
                    seenFirst = true;
                    previousHash = entry.Hash ?? "";
                    continue;
                }

                if (!string.IsNullOrEmpty(entry.Hash)
                    && string.Equals(entry.Hash, previousHash, StringComparison.Ordinal))
                {
                    continue;
                }

                return i;
            }

            return -1;
        }

        private sealed class UserCheckpointRow
        {
            public OperationLogEntry Entry { get; set; }
            public string Baseline { get; set; }
            public List<string> HiddenTimestamps { get; } = new List<string>();
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

        public static string BindCheckpointIdentity(CheckpointLiveTarget live, bool create, out string error)
        {
            error = null;
            if (live == null || live.Live == null)
            {
                error = "没有活动的文档窗口";
                return null;
            }

            string existing = EasyWriteDocIdStore.TrySplitOnSaveAs(live.Live);
            if (string.IsNullOrEmpty(existing))
            {
                existing = EasyWriteDocIdStore.TryRead(live.Live);
            }

            if (string.IsNullOrEmpty(existing))
            {
                if (!create)
                {
                    return null;
                }

                existing = Uuid7Generator.GenerateUuid7();
                if (!EasyWriteDocIdStore.TryWrite(live.Live, existing))
                {
                    error = "创建文档副本失败: 无法写入文档身份 EasyWriteDocId";
                    return null;
                }
            }

            DocumentCheckpointSession session = GetOrCreateSession(existing);
            ApplyLiveMetadata(
                session,
                live.Live,
                EasyWriteDocIdStore.TryReadFullName(live.Live),
                live.Adapter?.CopyExtension);
            return existing;
        }

        public static bool IsRoundMutating(string docId)
        {
            return !string.IsNullOrEmpty(docId) && RoundMutatingChannelByDocId.ContainsKey(docId);
        }

        public static void TrySaveUserCheckpointsForRound()
        {
            if (RoundMutatingChannelByDocId.Count == 0)
            {
                return;
            }

            var touches = new List<KeyValuePair<string, string>>(RoundMutatingChannelByDocId);
            foreach (KeyValuePair<string, string> touch in touches)
            {
                try
                {
                    TrySaveUserCheckpoint(touch.Key, touch.Value);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[Checkpoint] user checkpoint failed: " + ex.Message);
                }
                finally
                {
                    RoundMutatingChannelByDocId.Remove(touch.Key);
                    RoundBeforeUserTimestampByDocId.Remove(touch.Key);
                }
            }
        }

        public static object ListUserCheckpoints(string channelId)
        {
            if (!TryLiveForChannel(channelId, out CheckpointLiveTarget live, out string error))
            {
                return new { success = false, message = error };
            }

            string docId = BindCheckpointIdentity(live, create: false, out error);
            if (string.IsNullOrEmpty(docId))
            {
                return new { success = true, items = new object[0] };
            }

            DocumentCheckpointSession session = GetOrCreateSession(docId);
            int branch = session.CurrentBranch > 0 ? session.CurrentBranch : 1;
            var kept = new List<UserCheckpointRow>();
            string previousHash = null;
            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry.Kind != OperationLogEntryKind.Checkpoint
                    || !entry.UserRollback
                    || entry.Branch != branch)
                {
                    continue;
                }

                if (kept.Count > 0
                    && !string.IsNullOrEmpty(entry.Hash)
                    && string.Equals(entry.Hash, previousHash, StringComparison.Ordinal))
                {
                    kept[kept.Count - 1].HiddenTimestamps.Add(entry.Timestamp);
                    continue;
                }

                kept.Add(new UserCheckpointRow { Entry = entry });
                previousHash = entry.Hash ?? "";
            }

            bool originAssigned = false;
            foreach (UserCheckpointRow row in kept)
            {
                if (row.Entry.Operations.Count > 0)
                {
                    continue;
                }

                row.Baseline = originAssigned ? "before" : "origin";
                originAssigned = true;
            }

            var items = new List<object>();
            for (int i = kept.Count - 1; i >= 0; i--)
            {
                UserCheckpointRow row = kept[i];
                OperationLogEntry entry = row.Entry;
                var operations = new List<object>();
                foreach (CheckpointOperationRecord op in entry.Operations)
                {
                    operations.Add(new { tool = op.Tool ?? "", arguments = op.Arguments ?? "{}" });
                }

                string baseline = row.Baseline ?? "";

                bool current = entry.Branch == session.PositionBranch
                    && (string.Equals(entry.Timestamp, session.PositionTimestamp, StringComparison.Ordinal)
                        || row.HiddenTimestamps.Contains(session.PositionTimestamp));
                items.Add(new
                {
                    timestamp = entry.Timestamp,
                    hash = entry.Hash ?? "",
                    current,
                    operations,
                    baseline
                });
            }

            return new { success = true, items };
        }

        public static ToolResult RestoreUserCheckpoint(string channelId, string timestamp)
        {
            if (!TryLiveForChannel(channelId, out CheckpointLiveTarget live, out string error))
            {
                return new ToolResult { Success = false, Error = error };
            }

            string docId = BindCheckpointIdentity(live, create: false, out error);
            if (string.IsNullOrEmpty(docId))
            {
                return new ToolResult { Success = false, Error = "这份文档还没有历史记录" };
            }

            if (IsRoundMutating(docId))
            {
                return new ToolResult { Success = false, Error = RoundMutatingMessage };
            }

            live.DocUuid = docId;
            return RestoreCheckpoint(timestamp, live);
        }

        private static void TrySaveUserCheckpoint(string docId, string channelId)
        {
            DocumentCheckpointSession session = GetSession(docId) ?? CheckpointHistoryStore.TryLoad(docId);
            if (session == null)
            {
                return;
            }

            Sessions[docId] = session;
            int branch = session.CurrentBranch > 0 ? session.CurrentBranch : 1;
            int start = SpanStartIndex(session, branch);
            if (!HasOrdinaryCheckpointAfter(session, branch, start))
            {
                return;
            }

            if (!TryLiveForChannel(channelId, out CheckpointLiveTarget live, out _))
            {
                return;
            }

            live.DocUuid = docId;
            CheckpointCaret caret = null;
            try
            {
                caret = live.Adapter.CaptureCaret(live);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Checkpoint] user caret failed: " + ex.Message);
            }

            string dir = GetDocCheckpointDirectory(docId);
            Directory.CreateDirectory(dir);
            string timestamp = GenerateUniqueTimestamp(dir);
            string path = GetCheckpointFilePath(docId, timestamp, live.Adapter.CopyExtension);
            live.Adapter.SaveCopy(live, path);
            if (caret != null)
            {
                session.CaretByTimestamp[timestamp] = caret;
            }

            string hash = CheckpointContentHash.FromFile(path);
            AppendCheckpointEntry(session, timestamp, branch, userRollback: true, CollectOperations(session, branch, start), hash);
            session.PositionBranch = branch;
            session.PositionTimestamp = timestamp;
            ApplyLiveMetadata(session, live.Live, EasyWriteDocIdStore.TryReadFullName(live.Live), live.Adapter.CopyExtension);
            EvictOldestUser(session, docId);
            Persist(session);
        }

        private static bool TryLiveForChannel(string channelId, out CheckpointLiveTarget live, out string error)
        {
            live = null;
            error = null;
            if (string.IsNullOrEmpty(channelId)
                || !ChannelRegistry.TryGet(channelId, out IOperationChannel channel)
                || channel == null)
            {
                error = "找不到这份打开的文档";
                return false;
            }

            if (!CheckpointKindAdapters.TryFromChannel(channel, out live, out string kindError))
            {
                error = string.IsNullOrEmpty(kindError) ? "当前渠道不能回滚" : kindError;
                return false;
            }

            return true;
        }

        private static void MaybeFork(DocumentCheckpointSession session)
        {
            int branch = session.CurrentBranch > 0 ? session.CurrentBranch : 1;
            session.CurrentBranch = branch;
            string tip = LatestUserTimestamp(session, branch);
            if (string.IsNullOrEmpty(tip) || string.IsNullOrEmpty(session.PositionTimestamp))
            {
                return;
            }

            if (session.PositionBranch == branch
                && string.Equals(session.PositionTimestamp, tip, StringComparison.Ordinal))
            {
                return;
            }

            if (session.PositionBranch == branch && IsAfterTip(session, branch, tip, session.PositionTimestamp))
            {
                return;
            }

            int next = 1;
            foreach (CheckpointBranch item in session.Branches)
            {
                if (item.Id >= next)
                {
                    next = item.Id + 1;
                }
            }

            session.Branches.Add(new CheckpointBranch
            {
                Id = next,
                ForkedFromBranch = session.PositionBranch > 0 ? session.PositionBranch : branch,
                ForkedFromTimestamp = session.PositionTimestamp
            });
            session.CurrentBranch = next;
        }

        private static void RevertBranchIfEmpty(DocumentCheckpointSession session)
        {
            if (session == null || session.CurrentBranch <= 1)
            {
                return;
            }

            int branch = session.CurrentBranch;
            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry.Kind == OperationLogEntryKind.Checkpoint && entry.Branch == branch)
                {
                    return;
                }
            }

            CheckpointBranch info = null;
            foreach (CheckpointBranch item in session.Branches)
            {
                if (item.Id == branch)
                {
                    info = item;
                    break;
                }
            }

            if (info == null)
            {
                return;
            }

            session.Branches.Remove(info);
            if (info.ForkedFromBranch > 0)
            {
                session.CurrentBranch = info.ForkedFromBranch;
                session.PositionBranch = info.ForkedFromBranch;
                session.PositionTimestamp = info.ForkedFromTimestamp;
            }
        }

        private static int SpanStartIndex(DocumentCheckpointSession session, int branch)
        {
            int start = -1;
            for (int i = 0; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && entry.UserRollback
                    && entry.Branch == branch)
                {
                    start = i;
                }
            }

            return start;
        }

        private static bool HasOrdinaryCheckpointAfter(DocumentCheckpointSession session, int branch, int start)
        {
            for (int i = start + 1; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && !entry.UserRollback
                    && entry.Branch == branch)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<CheckpointOperationRecord> CollectOperations(
            DocumentCheckpointSession session,
            int branch,
            int start)
        {
            var operations = new List<CheckpointOperationRecord>();
            for (int i = start + 1; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind != OperationLogEntryKind.Action || entry.Branch != branch)
                {
                    continue;
                }

                operations.Add(new CheckpointOperationRecord
                {
                    Tool = entry.ToolName,
                    Arguments = entry.ActionLineBody
                });
            }

            return operations;
        }

        private static string LatestUserTimestamp(DocumentCheckpointSession session, int branch)
        {
            string timestamp = null;
            foreach (OperationLogEntry entry in session.Entries)
            {
                if (entry.Kind == OperationLogEntryKind.Checkpoint
                    && entry.UserRollback
                    && entry.Branch == branch)
                {
                    timestamp = entry.Timestamp;
                }
            }

            return timestamp;
        }

        private static bool IsAfterTip(
            DocumentCheckpointSession session,
            int branch,
            string tip,
            string timestamp)
        {
            int tipIndex = -1;
            int posIndex = -1;
            for (int i = 0; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry.Kind != OperationLogEntryKind.Checkpoint || entry.Branch != branch)
                {
                    continue;
                }

                if (string.Equals(entry.Timestamp, tip, StringComparison.Ordinal))
                {
                    tipIndex = i;
                }

                if (string.Equals(entry.Timestamp, timestamp, StringComparison.Ordinal))
                {
                    posIndex = i;
                }
            }

            return tipIndex >= 0 && posIndex > tipIndex;
        }

        private static void EvictOldestOrdinary(DocumentCheckpointSession session, string docId)
        {
            int branch = session.CurrentBranch > 0 ? session.CurrentBranch : 1;
            while (CountUsableOrdinary(session, branch) > MaxOrdinaryPerBranch)
            {
                int index = FindOldestUsableOrdinary(session, branch);
                if (index < 0)
                {
                    return;
                }

                RemoveOrdinaryEntry(session, docId, index);
            }
        }

        private static void EvictOldestUser(DocumentCheckpointSession session, string docId)
        {
            int branch = session.CurrentBranch > 0 ? session.CurrentBranch : 1;
            while (CountBillableUsers(session, branch) > MaxUserPerBranch)
            {
                int index = FindOldestBillableUser(session, branch);
                if (index < 0)
                {
                    return;
                }

                RemoveSpecificCheckpoint(session, docId, index);
            }
        }

        private static OperationLogEntry FindCheckpointEntry(DocumentCheckpointSession session, string timestamp)
        {
            int index = FindCheckpointEntryIndex(session, timestamp);
            return index < 0 ? null : session.Entries[index];
        }

        private static void ApplyLiveMetadata(DocumentCheckpointSession session, object live, string fullName, string extension = null)
        {
            if (session == null)
            {
                return;
            }

            string name = EasyWriteDocIdStore.TryReadName(live);
            if (!string.IsNullOrEmpty(name))
            {
                session.Name = name;
            }

            if (EasyWriteDocIdStore.IsPersistedPath(fullName))
            {
                session.Path = fullName;
                if (string.IsNullOrEmpty(extension))
                {
                    extension = Path.GetExtension(fullName);
                }
            }

            if (!string.IsNullOrEmpty(extension))
            {
                session.Kind = EasyWriteDocIdStore.KindFromExtension(extension.TrimStart('.'));
            }
        }

        private static void Persist(DocumentCheckpointSession session)
        {
            try
            {
                CheckpointHistoryStore.Save(session);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Checkpoint] history save failed: " + ex.Message);
            }
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
