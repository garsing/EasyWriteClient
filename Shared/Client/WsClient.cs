using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace WordAddIn1
{
    /// <summary>
    /// Word 插件 Client WebSocket（I6 启动即连 + bind）。
    /// </summary>
    public class WsClient : IDisposable
    {
        private readonly Control _syncControl;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer();
        private ClientWebSocket _socket;
        private CancellationTokenSource _cts;
        private WsMethodRegistry _registry;
        private object _wordApplication;
        private string _baseUrl;
        private string _connectedAccessToken;
        private string _connectedUsername;
        private bool _disposed;
        private int _reconnectDelayMs = 2000;
        private string _boundConversationId;
        private readonly object _bindLock = new object();
        private TaskCompletionSource<bool> _bindTcs;
        private string _pendingBindConversationId;
        private TaskCompletionSource<bool> _unbindTcs;
        private string _activeInvokeRequestId;
        private volatile bool _localInvokeCancelled;
        private readonly object _invokeCancelLock = new object();
        private readonly SemaphoreSlim _connectLock = new SemaphoreSlim(1, 1);
        private int _reconnectGeneration;

        public bool IsConnected =>
            _socket != null && _socket.State == WebSocketState.Open;

        public string BoundConversationId => _boundConversationId;

        public event Action<string> BindError;

        /// <summary>
        /// 服务端 UI 推送（如 todo_list_updated / request_cancel），由 TaskPane 转发 bridge。
        /// </summary>
        public event Action<string, Dictionary<string, object>> ServerUiMessage;

        /// <summary>invoke 完成后回传 tool_call_id / method / ToolResult，供对话气泡贴输出。</summary>
        public event Action<string, string, ToolResult> InvokeCompleted;

        public WsClient(Control syncControl)
        {
            _syncControl = syncControl ?? throw new ArgumentNullException(nameof(syncControl));
            // 截图 base64 回传可能较大（默认 MaxJsonLength 仅 2MB）
            _json.MaxJsonLength = 8 * 1024 * 1024;
        }

        private static readonly SemaphoreSlim InvokeGate = new SemaphoreSlim(1, 1);

        /// <summary>
        /// 注入 Word.Application。Desktop 启动时可为 null：仍须建立 Registry，
        /// 以便 document.open（F_open_document）按需 new Word；否则会误报 not ready。
        /// </summary>
        public void SetWordApplication(object wordApplication)
        {
            _wordApplication = wordApplication;
            _registry = new WsMethodRegistry(_wordApplication);
        }

        public async Task ConnectAsync(string httpBaseUrl, string accessToken, string username)
        {
            if (string.IsNullOrWhiteSpace(accessToken) || string.IsNullOrWhiteSpace(username))
            {
                return;
            }

            await _connectLock.WaitAsync().ConfigureAwait(false);
            try
            {
                _baseUrl = httpBaseUrl?.TrimEnd('/') ?? ConfigManager.Config.Api.BaseUrl.TrimEnd('/');

                if (IsConnected &&
                    string.Equals(_connectedAccessToken, accessToken, StringComparison.Ordinal) &&
                    string.Equals(_connectedUsername, username, StringComparison.Ordinal))
                {
                    return;
                }

                await DisconnectInternalAsync().ConfigureAwait(false);
                await EnsureConnectedAsync(accessToken, username).ConfigureAwait(false);
            }
            finally
            {
                _connectLock.Release();
            }
        }

        public async Task DisconnectAsync()
        {
            await _connectLock.WaitAsync().ConfigureAwait(false);
            try
            {
                await DisconnectInternalAsync().ConfigureAwait(false);
            }
            finally
            {
                _connectLock.Release();
            }
        }

        private async Task DisconnectInternalAsync()
        {
            Interlocked.Increment(ref _reconnectGeneration);
            _connectedAccessToken = null;
            _connectedUsername = null;
            await SafeCloseAndDisposeAsync().ConfigureAwait(false);
            EasyWriteDiagnostics.Log(DebugCategory.Ws, "[WsClient] disconnected");
        }

        public async Task<bool> BindAsync(string conversationId, int timeoutMs = 5000)
        {
            if (string.IsNullOrWhiteSpace(conversationId) || conversationId == "-1" || conversationId == "0")
            {
                return false;
            }

            if (!IsConnected)
            {
                return false;
            }

            lock (_bindLock)
            {
                if (_boundConversationId == conversationId)
                {
                    return true;
                }
            }

            TaskCompletionSource<bool> tcs;
            lock (_bindLock)
            {
                _pendingBindConversationId = conversationId;
                _bindTcs = new TaskCompletionSource<bool>();
                tcs = _bindTcs;
            }

            var payload = _json.Serialize(new Dictionary<string, object>
            {
                { "type", "bind" },
                { "conversation_id", conversationId }
            });
            await SendTextAsync(payload);

            var completed = await Task.WhenAny(tcs.Task, Task.Delay(timeoutMs));
            if (completed != tcs.Task)
            {
                lock (_bindLock)
                {
                    if (_bindTcs == tcs)
                    {
                        _bindTcs = null;
                        _pendingBindConversationId = null;
                    }
                }
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] bind timeout: {conversationId}");
                return false;
            }

            return await tcs.Task;
        }

        public async Task<bool> UnbindAsync(int timeoutMs = 3000)
        {
            lock (_bindLock)
            {
                if (string.IsNullOrEmpty(_boundConversationId))
                {
                    return true;
                }
            }

            if (!IsConnected)
            {
                lock (_bindLock)
                {
                    _boundConversationId = null;
                }
                return true;
            }

            TaskCompletionSource<bool> tcs;
            lock (_bindLock)
            {
                _unbindTcs = new TaskCompletionSource<bool>();
                tcs = _unbindTcs;
            }

            var payload = _json.Serialize(new Dictionary<string, object>
            {
                { "type", "unbind" }
            });
            await SendTextAsync(payload);

            var completed = await Task.WhenAny(tcs.Task, Task.Delay(timeoutMs));
            if (completed != tcs.Task)
            {
                lock (_bindLock)
                {
                    if (_unbindTcs == tcs)
                    {
                        _unbindTcs = null;
                    }
                    _boundConversationId = null;
                }
                EasyWriteDiagnostics.Log(DebugCategory.Ws, "[WsClient] unbind timeout");
                return true;
            }

            return await tcs.Task;
        }

        public async Task SendRunCancelAsync(string conversationId)
        {
            if (string.IsNullOrWhiteSpace(conversationId) || conversationId == "-1" || conversationId == "0")
            {
                return;
            }

            if (!IsConnected)
            {
                return;
            }

            var payload = _json.Serialize(new Dictionary<string, object>
            {
                { "type", "run_cancel" },
                { "conversation_id", conversationId }
            });
            try
            {
                await SendTextAsync(payload);
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] SendRunCancelAsync failed: {ex.Message}");
            }
        }

        public void CancelActiveInvokeLocal()
        {
            lock (_invokeCancelLock)
            {
                _localInvokeCancelled = true;
            }
        }

        private bool ShouldCancelInvoke(string requestId)
        {
            lock (_invokeCancelLock)
            {
                if (!_localInvokeCancelled)
                {
                    return AgentRunCancellation.Token.IsCancellationRequested;
                }
                return string.IsNullOrEmpty(_activeInvokeRequestId) || _activeInvokeRequestId == requestId;
            }
        }

        private void BeginInvokeTracking(string requestId)
        {
            lock (_invokeCancelLock)
            {
                _activeInvokeRequestId = requestId;
                _localInvokeCancelled = false;
            }
        }

        private void EndInvokeTracking(string requestId)
        {
            lock (_invokeCancelLock)
            {
                if (_activeInvokeRequestId == requestId)
                {
                    _activeInvokeRequestId = null;
                }
                _localInvokeCancelled = false;
            }
        }

        private async Task EnsureConnectedAsync(string accessToken, string username)
        {
            await SafeCloseAndDisposeAsync().ConfigureAwait(false);

            _socket = new ClientWebSocket();
            _cts = new CancellationTokenSource();
            var generation = _reconnectGeneration;

            var wsUrl = BuildWsUrl(_baseUrl, accessToken, username);
            try
            {
                await _socket.ConnectAsync(new Uri(wsUrl), _cts.Token);
                _connectedAccessToken = accessToken;
                _connectedUsername = username;
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] connected: {wsUrl.Split('?')[0]} user={username}");

                var methods = _registry?.GetMethodNames() ?? new List<string>();
                var connectedMsg = _json.Serialize(new Dictionary<string, object>
                {
                    { "type", "connected" },
                    { "methods", methods }
                });
                await SendTextAsync(connectedMsg);

                _ = Task.Run(() => ReceiveLoopAsync(_cts.Token));
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] connect failed: {ex.Message}");
                await SafeCloseAndDisposeAsync().ConfigureAwait(false);
                ScheduleReconnect(accessToken, username, generation);
            }
        }

        private string BuildWsUrl(string httpBaseUrl, string token, string username)
        {
            var uri = new Uri(httpBaseUrl);
            var scheme = uri.Scheme == "https" ? "wss" : "ws";
            var encodedToken = Uri.EscapeDataString(token);
            var encodedUser = Uri.EscapeDataString(username);
            return $"{scheme}://{uri.Authority}/api/client/ws?token={encodedToken}&username={encodedUser}";
        }

        private async Task ReceiveLoopAsync(CancellationToken token)
        {
            var buffer = new ArraySegment<byte>(new byte[8192]);
            var builder = new StringBuilder();

            while (!token.IsCancellationRequested && _socket != null && _socket.State == WebSocketState.Open)
            {
                builder.Clear();
                WebSocketReceiveResult result;
                do
                {
                    result = await _socket.ReceiveAsync(buffer, token);
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        await SafeCloseAndDisposeAsync();
                        return;
                    }
                    builder.Append(Encoding.UTF8.GetString(buffer.Array, 0, result.Count));
                } while (!result.EndOfMessage);

                var text = builder.ToString();
                if (text == "pong")
                {
                    continue;
                }

                try
                {
                    var msg = _json.Deserialize<Dictionary<string, object>>(text);
                    await HandleMessageAsync(msg);
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] message parse error: {ex.Message}");
                }
            }
        }

        private async Task HandleMessageAsync(Dictionary<string, object> msg)
        {
            if (msg == null || !msg.TryGetValue("type", out var typeObj))
            {
                return;
            }

            var type = typeObj as string;
            if (type == "bind_error")
            {
                var err = msg.ContainsKey("error") ? msg["error"]?.ToString() : "bind failed";
                BindError?.Invoke(err);
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] bind_error: {err}");
                CompletePendingBind(false);
                return;
            }

            if (type == "bind_ok")
            {
                var conv = msg.ContainsKey("conversation_id") ? msg["conversation_id"]?.ToString() : "?";
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] bind_ok: {conv}");
                lock (_bindLock)
                {
                    _boundConversationId = conv;
                }
                SyncConversationContextFromBound();
                CompletePendingBind(true);
                return;
            }

            if (type == "unbind_ok")
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, "[WsClient] unbind_ok");
                lock (_bindLock)
                {
                    _boundConversationId = null;
                }
                CompletePendingUnbind(true);
                return;
            }

            if (type == "unbind_error")
            {
                var err = msg.ContainsKey("error") ? msg["error"]?.ToString() : "unbind failed";
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] unbind_error: {err}");
                lock (_bindLock)
                {
                    _boundConversationId = null;
                }
                CompletePendingUnbind(true);
                return;
            }

            if (type == "invoke_cancel" || type == "request_cancel")
            {
                var rid = msg.ContainsKey("request_id") ? msg["request_id"]?.ToString() : null;
                lock (_invokeCancelLock)
                {
                    if (string.IsNullOrEmpty(rid) || rid == _activeInvokeRequestId)
                    {
                        _localInvokeCancelled = true;
                    }
                }
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] {type} request_id={rid ?? "*"}");
                try
                {
                    ServerUiMessage?.Invoke(type, msg);
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] ServerUiMessage cancel err: {ex.Message}");
                }
                return;
            }

            if (type == "todo_list_updated")
            {
                // 勿阻塞 ReceiveLoop：交给宿主 / Vue 处理后 SendReply。
                try
                {
                    ServerUiMessage?.Invoke(type, msg);
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] ServerUiMessage todo err: {ex.Message}");
                    var rid = msg.ContainsKey("request_id") ? msg["request_id"]?.ToString() : null;
                    if (!string.IsNullOrEmpty(rid))
                    {
                        _ = SendReplyAsync(rid, false, null, ex.Message);
                    }
                }
                return;
            }

            if (type == "title_updated")
            {
                // notify：无需 reply；Desktop 侧栏就地更新标题。
                try
                {
                    ServerUiMessage?.Invoke(type, msg);
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] ServerUiMessage title err: {ex.Message}");
                }
                return;
            }

            if (type == "invoke")
            {
                // 勿 await：长耗时工具会阻塞 ReceiveAsync，导致 ping 超时断连。
                _ = HandleInvokeAsyncSafe(msg);
                return;
            }
        }

        private async Task HandleInvokeAsyncSafe(Dictionary<string, object> msg)
        {
            await InvokeGate.WaitAsync().ConfigureAwait(false);
            try
            {
                await HandleInvokeAsync(msg).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] HandleInvokeAsync unhandled: {ex}");
            }
            finally
            {
                InvokeGate.Release();
            }
        }

        private async Task HandleInvokeAsync(Dictionary<string, object> msg)
        {
            var requestId = msg.ContainsKey("request_id") ? msg["request_id"]?.ToString() : null;
            var toolCallId = msg.ContainsKey("tool_call_id") ? msg["tool_call_id"]?.ToString() : null;
            var method = msg.ContainsKey("method") ? msg["method"]?.ToString() : null;
            var parameters = ExtractParams(msg);

            if (string.IsNullOrEmpty(requestId) || string.IsNullOrEmpty(toolCallId))
            {
                return;
            }

            BeginInvokeTracking(requestId);
            var invokeSw = Stopwatch.StartNew();
            SyncConversationContextFromBound();
            EasyWriteDiagnostics.Log(DebugCategory.Ws, 
                $"[WsClient] invoke START method={method} request_id={requestId} thread={Environment.CurrentManagedThreadId} conv={ConversationContext.CurrentId}");

            try
            {
                AgentRunCancellation.ThrowIfCancelled();
                if (ShouldCancelInvoke(requestId))
                {
                    await SendInvokeResultAsync(requestId, toolCallId, false, null, "cancelled by user");
                    return;
                }

                if (_registry == null)
                {
                    await SendInvokeResultAsync(requestId, toolCallId, false, null, "Word application not ready");
                    return;
                }

                ToolResult toolResult = await RunInvokeOnAffinityAsync(method, () =>
                    _registry.InvokeRawAsync(method, parameters));

                if (ShouldCancelInvoke(requestId))
                {
                    await SendInvokeResultAsync(requestId, toolCallId, false, null, "cancelled by user");
                    return;
                }

                bool ok = toolResult.Success;
                string resultText = _json.Serialize(new
                {
                    success = toolResult.Success,
                    message = toolResult.Success ? "工具执行成功" : toolResult.Error,
                    data = toolResult.Data
                });
                string errorText = null;
                if (!ok)
                {
                    errorText = toolResult.Error ?? resultText;
                }
                TryRaiseInvokeCompleted(toolCallId, method, toolResult);
                await SendInvokeResultAsync(requestId, toolCallId, ok, resultText, errorText);
                EasyWriteDiagnostics.Log(DebugCategory.Ws, 
                    $"[WsClient] invoke END method={method} ok={ok} elapsed_ms={invokeSw.ElapsedMilliseconds} result_len={resultText?.Length ?? 0}");
                EasyWriteDiagnostics.LogTiming(
                    "ws.invoke",
                    invokeSw.ElapsedMilliseconds,
                    $"method={method} ok={ok}");
                if (!ok && !string.IsNullOrEmpty(errorText))
                {
                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[WsClient] invoke error method={method} request_id={requestId} error={errorText}");
                }
            }
            catch (OperationCanceledException)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, 
                    $"[WsClient] invoke CANCELLED method={method} elapsed_ms={invokeSw.ElapsedMilliseconds}");
                EasyWriteDiagnostics.LogTiming(
                    "ws.invoke",
                    invokeSw.ElapsedMilliseconds,
                    $"method={method} ok=cancelled");
                await SendInvokeResultAsync(requestId, toolCallId, false, null, "cancelled by user");
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, 
                    $"[WsClient] invoke FAIL method={method} elapsed_ms={invokeSw.ElapsedMilliseconds} err={ex.Message}");
                EasyWriteDiagnostics.LogTiming(
                    "ws.invoke",
                    invokeSw.ElapsedMilliseconds,
                    $"method={method} ok=fail");
                await SendInvokeResultAsync(requestId, toolCallId, false, null, ex.Message);
            }
            finally
            {
                EndInvokeTracking(requestId);
            }
        }

        /// <summary>
        /// Desktop：Office COM 走专属 STA；browser.* 走主窗 UI。
        /// Plugin：专属 STA 未启动，整段仍走任务窗格线程（Word 主 STA）。
        /// </summary>
        private Task<T> RunInvokeOnAffinityAsync<T>(string method, Func<Task<T>> action)
        {
            if (OfficeStaScheduler.IsEnabled && !NeedsUiThread(method))
            {
                return OfficeStaScheduler.InvokeAsync(action);
            }

            return RunOnUiThreadAsync(action);
        }

        private static bool NeedsUiThread(string method)
        {
            return !string.IsNullOrEmpty(method)
                && method.StartsWith("browser.", StringComparison.OrdinalIgnoreCase);
        }

        private Task<T> RunOnUiThreadAsync<T>(Func<Task<T>> action)
        {
            if (!_syncControl.InvokeRequired)
            {
                return action();
            }

            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            _syncControl.BeginInvoke(new Action(async () =>
            {
                try
                {
                    tcs.TrySetResult(await action().ConfigureAwait(true));
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }));
            return tcs.Task;
        }

        private static Dictionary<string, object> ExtractParams(Dictionary<string, object> msg)
        {
            if (!msg.ContainsKey("params") || msg["params"] == null)
            {
                return new Dictionary<string, object>();
            }

            var raw = msg["params"];
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            try
            {
                var serializer = new JavaScriptSerializer();
                var json = serializer.Serialize(raw);
                return serializer.Deserialize<Dictionary<string, object>>(json)
                       ?? new Dictionary<string, object>();
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }

        private async Task SendInvokeResultAsync(
            string requestId,
            string toolCallId,
            bool ok,
            string result,
            string error)
        {
            var payload = new Dictionary<string, object>
            {
                { "type", "invoke_result" },
                { "request_id", requestId },
                { "tool_call_id", toolCallId },
                { "ok", ok }
            };
            if (ok)
            {
                payload["result"] = result;
            }
            else
            {
                payload["error"] = error ?? "invoke failed";
            }
            await SendTextAsync(_json.Serialize(payload));
        }

        /// <summary>
        /// 通用 request 回包（Todo UI ack 等）。
        /// </summary>
        public async Task SendReplyAsync(string requestId, bool ok, object data = null, string error = null)
        {
            if (string.IsNullOrEmpty(requestId))
            {
                return;
            }
            var payload = new Dictionary<string, object>
            {
                { "type", "reply" },
                { "request_id", requestId },
                { "ok", ok }
            };
            if (ok)
            {
                if (data != null)
                {
                    payload["data"] = data;
                }
            }
            else
            {
                payload["error"] = error ?? "request failed";
            }
            await SendTextAsync(_json.Serialize(payload));
        }

        private async Task SendTextAsync(string text)
        {
            if (_socket == null || _socket.State != WebSocketState.Open)
            {
                return;
            }
            var bytes = Encoding.UTF8.GetBytes(text);
            await _socket.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, _cts.Token);
        }

        private void ScheduleReconnect(string accessToken, string username, int generation)
        {
            _ = Task.Run(async () =>
            {
                await Task.Delay(_reconnectDelayMs);
                if (_disposed || generation != _reconnectGeneration)
                {
                    return;
                }
                if (!UserService.Instance.CheckLoginStatus())
                {
                    return;
                }
                var user = UserService.Instance;
                if (!string.Equals(user.AccessToken, accessToken, StringComparison.Ordinal) ||
                    !string.Equals(user.UserName, username, StringComparison.Ordinal))
                {
                    return;
                }
                try
                {
                    await ConnectAsync(_baseUrl, accessToken, username).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Ws, $"[WsClient] reconnect failed: {ex.Message}");
                }
            });
        }

        private async Task SafeCloseAndDisposeAsync()
        {
            var socket = _socket;
            var cts = _cts;
            _socket = null;
            _cts = null;

            try
            {
                cts?.Cancel();
            }
            catch
            {
                // ignore
            }

            try
            {
                if (socket != null && socket.State == WebSocketState.Open)
                {
                    await socket.CloseAsync(WebSocketCloseStatus.NormalClosure, "closing", CancellationToken.None);
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                socket?.Dispose();
            }
            catch
            {
                // ignore
            }

            try
            {
                cts?.Dispose();
            }
            catch
            {
                // ignore
            }

            lock (_bindLock)
            {
                _boundConversationId = null;
                _bindTcs?.TrySetResult(false);
                _bindTcs = null;
                _pendingBindConversationId = null;
                _unbindTcs?.TrySetResult(false);
                _unbindTcs = null;
            }
        }

        private void CompletePendingBind(bool success)
        {
            TaskCompletionSource<bool> tcs;
            lock (_bindLock)
            {
                tcs = _bindTcs;
                _bindTcs = null;
                _pendingBindConversationId = null;
            }
            tcs?.TrySetResult(success);
        }

        /// <summary>
        /// F_* 工具读工作区文件时用 ConversationContext；以 WS bind 的会话 ID 为准，避免仍为 -1 落到 sessions/_pending。
        /// </summary>
        private void SyncConversationContextFromBound()
        {
            string bound;
            lock (_bindLock)
            {
                bound = _boundConversationId;
            }

            if (string.IsNullOrWhiteSpace(bound) || bound == "-1" || bound == "0")
            {
                return;
            }

            if (ConversationContext.CurrentId != bound)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, 
                    $"[WsClient] SyncConversationContext: {ConversationContext.CurrentId} -> {bound}");
                ConversationContext.CurrentId = bound;
            }
        }

        private void CompletePendingUnbind(bool success)
        {
            TaskCompletionSource<bool> tcs;
            lock (_bindLock)
            {
                tcs = _unbindTcs;
                _unbindTcs = null;
            }
            tcs?.TrySetResult(success);
        }

        private void DisposeSocketOnly()
        {
            try
            {
                _cts?.Cancel();
            }
            catch
            {
                // ignore
            }

            try
            {
                _socket?.Dispose();
            }
            catch
            {
                // ignore
            }

            _socket = null;
            _cts = null;
            lock (_bindLock)
            {
                _boundConversationId = null;
                _bindTcs?.TrySetResult(false);
                _bindTcs = null;
                _pendingBindConversationId = null;
                _unbindTcs?.TrySetResult(false);
                _unbindTcs = null;
            }
        }

        private void TryRaiseInvokeCompleted(string toolCallId, string method, ToolResult toolResult)
        {
            if (string.IsNullOrEmpty(toolCallId) || InvokeCompleted == null)
            {
                return;
            }

            try
            {
                InvokeCompleted(toolCallId, method ?? "", toolResult);
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Ws, "[WsClient] InvokeCompleted: " + ex.Message);
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            DisposeSocketOnly();
        }
    }
}
