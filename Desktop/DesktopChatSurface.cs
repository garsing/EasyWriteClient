using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// Desktop 主区：WebView2 + Vue（host=desktop）+ 聊天/登录/会话桥接（B4）。
    /// </summary>
    internal sealed class DesktopChatSurface : UserControl
    {
        private const string ApiKey = "sk-2e2929af6bde429990eef75c39e6afd7";

        private WebView2 _webView;
        private WebView2Bridge _bridge;
        private McpClient _mcpClient;
        private WsClient _wsClient;
        private readonly List<ChatMessage> _conversationHistory = new List<ChatMessage>();
        private string _currentConversationId = "-1";
        private System.Threading.CancellationTokenSource _cts;
        private string _authoritativeUploadId;
        private bool _uploadUiNotified;
        private bool _isProcessing;
        private volatile bool _newSessionResetPending;
        private bool _webReady;

        public DesktopChatSurface()
        {
            Dock = DockStyle.Fill;
            _webView = new WebView2
            {
                Dock = DockStyle.Fill,
                DefaultBackgroundColor = System.Drawing.Color.FromArgb(0xF0, 0xF0, 0xF0)
            };
            Controls.Add(_webView);
            UserService.Instance.OnUserLoggedIn += OnUserLoggedIn;
            UserService.Instance.OnUserLoggedOut += OnUserLoggedOut;
            HostCallbacks.WordApplicationResolved = OnWordApplicationResolved;
            Disposed += (_, __) =>
            {
                UserService.Instance.OnUserLoggedIn -= OnUserLoggedIn;
                UserService.Instance.OnUserLoggedOut -= OnUserLoggedOut;
                if (_wsClient != null)
                {
                    _wsClient.ServerUiMessage -= OnWsServerUiMessage;
                }
                if (ReferenceEquals(HostCallbacks.WordApplicationResolved, (Action<object>)OnWordApplicationResolved))
                {
                    HostCallbacks.WordApplicationResolved = null;
                }
            };
        }

        private void OnWordApplicationResolved(object wordApp)
        {
            WordHost.Attach(wordApp);
            if (_wsClient != null)
            {
                _wsClient.SetWordApplication(wordApp);
            }
        }

        public async Task InitializeAsync()
        {
            if (_webReady)
            {
                return;
            }

            // WebView2 在 Size=0 时初始化容易一直白屏；先走 Dock 布局拿到「顶栏下方」区域
            // 切勿把 Height 设成父窗 ClientSize.Height，否则会铺满整窗并被顶栏盖住内容
            if (Width < 32 || Height < 32)
            {
                var parent = FindForm();
                if (parent != null)
                {
                    parent.PerformLayout();
                    if (Width < 32 || Height < 32)
                    {
                        int topDock = 0;
                        foreach (Control c in parent.Controls)
                        {
                            if (c != this && c.Dock == DockStyle.Top)
                            {
                                topDock += c.Height;
                            }
                        }

                        Bounds = new System.Drawing.Rectangle(
                            0,
                            topDock,
                            Math.Max(parent.ClientSize.Width, 800),
                            Math.Max(parent.ClientSize.Height - topDock, 600));
                    }
                }
            }

            string logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EasyWriteDesktop");
            Directory.CreateDirectory(logDir);
            string logPath = Path.Combine(logDir, "webview-init.log");
            void Log(string msg)
            {
                string line = DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg;
                System.Diagnostics.Debug.WriteLine("[DesktopChatSurface] " + line);
                try
                {
                    File.AppendAllText(logPath, line + Environment.NewLine);
                }
                catch
                {
                }
            }

            Log($"begin Size={Size} ClientSize={ClientSize} www check…");

            string userData = Path.Combine(logDir, "WebView2Data");
            Directory.CreateDirectory(userData);

            var env = await CoreWebView2Environment.CreateAsync(userDataFolder: userData)
                .ConfigureAwait(true);
            await _webView.EnsureCoreWebView2Async(env).ConfigureAwait(true);

            string wwwroot = ResolveWwwroot();
            if (!Directory.Exists(wwwroot) || !File.Exists(Path.Combine(wwwroot, "index.html")))
            {
                throw new DirectoryNotFoundException("未找到前端 wwwroot/index.html: " + wwwroot);
            }

            string mainJs = Path.Combine(wwwroot, "assets", "main.js");
            Log("wwwroot=" + wwwroot + " main.js=" + File.Exists(mainJs));

            var core = _webView.CoreWebView2;
            core.Settings.AreDevToolsEnabled = true;
            core.Settings.AreDefaultScriptDialogsEnabled = true;
            core.Settings.IsScriptEnabled = true;
            core.Settings.IsWebMessageEnabled = true;
            core.Settings.AreDefaultContextMenusEnabled = true;

            core.SetVirtualHostNameToFolderMapping(
                "appassets.local",
                wwwroot,
                CoreWebView2HostResourceAccessKind.Allow);

            core.ProcessFailed += (_, e) =>
                Log("ProcessFailed kind=" + e.ProcessFailedKind);

            var navTcs = new TaskCompletionSource<bool>();
            core.NavigationCompleted += (_, e) =>
            {
                Log("NavigationCompleted success=" + e.IsSuccess
                    + " status=" + e.WebErrorStatus
                    + " url=" + core.Source);
                navTcs.TrySetResult(e.IsSuccess);
            };

            // 捕获前端脚本错误到日志
            await core.AddScriptToExecuteOnDocumentCreatedAsync(
                @"(function(){
  window.addEventListener('error', function(e){
    console.error('[EasyWrite] error', e.message, e.filename, e.lineno);
  });
  window.addEventListener('unhandledrejection', function(e){
    console.error('[EasyWrite] unhandledrejection', e.reason);
  });
})();").ConfigureAwait(true);

            _bridge = new WebView2Bridge(_webView);
            RegisterHandlers();

            // 先加载前端；勿在启动时 new Word
            const string url = "http://appassets.local/index.html?host=desktop";
            Log("Navigate " + url);
            core.Navigate(url);

            // 强制一次尺寸刷新（部分机器上 WebView2 初始化后需 resize 才绘制）
            var sz = _webView.Size;
            if (sz.Width > 0 && sz.Height > 0)
            {
                _webView.Size = new System.Drawing.Size(sz.Width + 1, sz.Height);
                _webView.Size = sz;
            }

            var completed = await Task.WhenAny(navTcs.Task, Task.Delay(15000)).ConfigureAwait(true);
            if (completed != navTcs.Task)
            {
                Log("NavigationCompleted timeout");
                MessageBox.Show(
                    "前端页面加载超时。\r\n日志: " + logPath,
                    "易写",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            else if (!navTcs.Task.Result)
            {
                MessageBox.Show(
                    "前端页面导航失败。\r\n日志: " + logPath,
                    "易写",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                // 等模块/Vue 挂载；失败则打开 DevTools 并给出可见提示
                await Task.Delay(800).ConfigureAwait(true);
                try
                {
                    string check = await core.ExecuteScriptAsync(
                        @"(function(){
  var app=document.getElementById('app');
  return JSON.stringify({
    href: location.href,
    appHtmlLen: app ? app.innerHTML.length : -1,
    readyState: document.readyState,
    title: document.title
  });
})();").ConfigureAwait(true);
                    Log("vue-check " + check);
                    // ExecuteScriptAsync 返回 JSON 字符串字面量
                    if (check != null && check.Contains("\"appHtmlLen\":0"))
                    {
                        Log("Vue #app still empty — opening DevTools");
                        core.OpenDevToolsWindow();
                        MessageBox.Show(
                            "页面已打开但 Vue 未渲染（#app 为空）。\r\n已打开开发者工具，请查看 Console。\r\n日志: "
                            + logPath,
                            "易写",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    Log("vue-check failed: " + ex.Message);
                }
            }

            _webReady = true;

            try
            {
                EnsureClients(createWordIfMissing: false);
            }
            catch (Exception ex)
            {
                Log("EnsureClients(startup): " + ex.Message);
            }
        }

        public async Task EnsureLoggedInAsync()
        {
            var owner = FindForm() ?? (IWin32Window)this;

            if (!UserService.Instance.CheckLoginStatus())
            {
                await LoginForm.ShowDialogAsync(owner, logoutFirst: false)
                    .ConfigureAwait(true);
            }

            if (!UserService.Instance.CheckLoginStatus())
            {
                return;
            }

            // 对齐 Plugin ThisAddIn：已登录也必须拉工作区根，否则 F_* Ensure 报「工作区未初始化」
            try
            {
                await UserService.Instance.InitializeWorkspaceSettingsAsync(owner)
                    .ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[DesktopChatSurface] InitializeWorkspaceSettingsAsync: " + ex.Message);
            }

            await EnsureWsReadyAsync().ConfigureAwait(true);
        }

        private static string ResolveWwwroot()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string local = Path.GetFullPath(Path.Combine(baseDir, "wwwroot"));
            // 开发：相对 Desktop 输出目录回退到 Shared/Frontend/wwwroot
            string sharedWww = Path.GetFullPath(Path.Combine(
                baseDir, "..", "..", "..", "..", "Shared", "Frontend", "wwwroot"));

            bool localOk = Directory.Exists(local) && File.Exists(Path.Combine(local, "index.html"));
            bool sharedOk = Directory.Exists(sharedWww) && File.Exists(Path.Combine(sharedWww, "index.html"));

            // npm run build 更新 Shared/Frontend/wwwroot；两侧都存在时取较新的一份
            if (localOk && sharedOk)
            {
                DateTime localStamp = WwwrootStamp(local);
                DateTime sharedStamp = WwwrootStamp(sharedWww);
                return sharedStamp >= localStamp ? sharedWww : local;
            }

            if (sharedOk)
            {
                return sharedWww;
            }

            return local;
        }

        private static DateTime WwwrootStamp(string wwwroot)
        {
            string mainJs = Path.Combine(wwwroot, "assets", "main.js");
            if (File.Exists(mainJs))
            {
                return File.GetLastWriteTimeUtc(mainJs);
            }

            string index = Path.Combine(wwwroot, "index.html");
            return File.Exists(index) ? File.GetLastWriteTimeUtc(index) : DateTime.MinValue;
        }

        private void RegisterHandlers()
        {
            _bridge.RegisterHandler("sendMessage", HandleSendMessageAsync);
            _bridge.RegisterHandler("stopRequest", HandleStopRequestAsync);
            _bridge.RegisterHandler("openLoginWindow", async _ =>
            {
                await LoginForm.ShowDialogAsync(FindForm() ?? (IWin32Window)this, logoutFirst: true)
                    .ConfigureAwait(true);
                return new { success = true };
            });
            _bridge.RegisterHandler("addConversation", HandleAddConversationAsync);
            _bridge.RegisterHandler("openConversation", HandleOpenConversationAsync);
            _bridge.RegisterHandler("getApiConfig", data =>
                Task.FromResult(KnowledgeBaseService.GetApiConfig(data)));
            _bridge.RegisterHandler("getDocumentEmptyState", _ =>
                Task.FromResult<object>(new
                {
                    success = true,
                    data = new { documentEmpty = true }
                }));
            _bridge.RegisterHandler("getToolAlias", HandleGetToolAliasAsync);
            _bridge.RegisterHandler("openSettings", async _ =>
            {
                try
                {
                    if (!UserService.Instance.IsLoggedIn)
                    {
                        await LoginForm.ShowDialogAsync(FindForm() ?? (IWin32Window)this, logoutFirst: false)
                            .ConfigureAwait(true);
                        if (!UserService.Instance.CheckLoginStatus())
                        {
                            return new { success = false, message = "未登录" };
                        }
                    }

                    await UserSettingsForm.ShowAsync(FindForm() ?? (IWin32Window)this).ConfigureAwait(true);
                    return new { success = true };
                }
                catch (Exception ex)
                {
                    return new { success = false, message = ex.Message };
                }
            });
            _bridge.RegisterHandler("openKnowledgeBase", async _ =>
            {
                await Task.CompletedTask;
                return new { success = true, message = "桌面版知识库入口后续接入" };
            });
            _bridge.RegisterHandler("todoListReply", HandleTodoListReplyAsync);
        }

        /// <param name="createWordIfMissing">仅工具路径为 true；启动/登录为 false。</param>
        private void EnsureClients(bool createWordIfMissing = false)
        {
            object wordApp = createWordIfMissing
                ? WordHost.GetWordApplicationForTools()
                : WordHost.TryGetExistingWordApplication();

            // wordApp 可为 null：纯聊天不强制 Word；但必须挂上 WS Registry，
            // 否则 document.open 在进工具前就被 WsClient 以 not ready 拒绝。
            _mcpClient = new McpClient(ApiKey, wordApp);
            if (_wsClient == null)
            {
                _wsClient = new WsClient(this);
                _wsClient.ServerUiMessage += OnWsServerUiMessage;
            }

            _wsClient.SetWordApplication(wordApp);
        }

        private void OnWsServerUiMessage(string type, Dictionary<string, object> msg)
        {
            if (_bridge == null || msg == null)
            {
                return;
            }

            try
            {
                if (type == "todo_list_updated")
                {
                    var requestId = msg.ContainsKey("request_id") ? msg["request_id"]?.ToString() : null;
                    var version = msg.ContainsKey("version") ? msg["version"] : null;
                    var conversationId = msg.ContainsKey("conversation_id")
                        ? msg["conversation_id"]?.ToString()
                        : null;
                    object todos = new object[0];
                    if (msg.ContainsKey("todos") && msg["todos"] != null)
                    {
                        try
                        {
                            var todosJson = JsonConvert.SerializeObject(msg["todos"]);
                            todos = JsonConvert.DeserializeObject(todosJson) ?? new object[0];
                        }
                        catch
                        {
                            todos = msg["todos"];
                        }
                    }

                    void SendTodo()
                    {
                        _bridge.SendToJavaScript("todoListUpdated", new
                        {
                            requestId,
                            conversationId,
                            version,
                            todos
                        });
                    }

                    if (InvokeRequired)
                    {
                        BeginInvoke((MethodInvoker)SendTodo);
                    }
                    else
                    {
                        SendTodo();
                    }

                    return;
                }

                if (type == "request_cancel" || type == "invoke_cancel")
                {
                    var requestId = msg.ContainsKey("request_id") ? msg["request_id"]?.ToString() : null;
                    if (!string.IsNullOrEmpty(requestId))
                    {
                        void SendRevert()
                        {
                            _bridge.SendToJavaScript("todoListRevert", new { requestId });
                        }

                        if (InvokeRequired)
                        {
                            BeginInvoke((MethodInvoker)SendRevert);
                        }
                        else
                        {
                            SendRevert();
                        }
                    }

                    return;
                }

                if (type == "title_updated")
                {
                    var conversationId = msg.ContainsKey("conversation_id")
                        ? msg["conversation_id"]?.ToString()
                        : null;
                    var title = msg.ContainsKey("title") ? msg["title"]?.ToString() : null;
                    if (string.IsNullOrEmpty(conversationId) || string.IsNullOrEmpty(title))
                    {
                        return;
                    }

                    void SendTitle()
                    {
                        _bridge.SendToJavaScript("conversationTitleUpdated", new
                        {
                            conversationId,
                            title
                        });
                    }

                    if (InvokeRequired)
                    {
                        BeginInvoke((MethodInvoker)SendTitle);
                    }
                    else
                    {
                        SendTitle();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DesktopChatSurface] OnWsServerUiMessage: {ex.Message}");
            }
        }

        private void SyncConversationContext()
        {
            ConversationContext.CurrentId = _currentConversationId;
        }

        private async Task<object> HandleSendMessageAsync(object data)
        {
            try
            {
                var messageData = data as JObject
                    ?? JObject.Parse(data?.ToString() ?? "{}");
                string content = messageData["content"]?.ToString();
                if (string.IsNullOrEmpty(content))
                {
                    return new { success = false, message = "消息内容为空" };
                }

                if (!UserService.Instance.CheckLoginStatus())
                {
                    return new
                    {
                        success = false,
                        requiresLogin = true,
                        message = "用户未登录，请先登录",
                        userInput = content
                    };
                }

                // 发消息：附着已有 Word 即可，不主动 new（Agent 调打开文档工具时再创建）
                EnsureClients(createWordIfMissing: false);
                Invoke((MethodInvoker)delegate
                {
                    _cts = new System.Threading.CancellationTokenSource();
                    AgentRunCancellation.BeginRun(_cts);
                    _isProcessing = true;
                    _bridge.SendToJavaScript("requestStateChanged", new { isProcessing = true });
                });

                _conversationHistory.Add(new ChatMessage { role = "user", content = content });
                _bridge.SendToJavaScript("userMessage", new
                {
                    id = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                    content,
                    timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds()
                });

                string uploadId = messageData["uploadId"]?.ToString()
                    ?? messageData["upload_id"]?.ToString();
                if (string.IsNullOrWhiteSpace(uploadId))
                {
                    uploadId = null;
                }

                _authoritativeUploadId = uploadId;
                _uploadUiNotified = false;
                await ProcessChatRequestAsync(content).ConfigureAwait(true);
                return new { success = true };
            }
            catch (Exception ex)
            {
                return new { success = false, message = ex.Message };
            }
        }

        private async Task ProcessChatRequestAsync(string userInput)
        {
            var responseBuilder = new System.Text.StringBuilder();
            var reasoningBuilder = new System.Text.StringBuilder();
            long messageId = DateTimeOffset.Now.ToUnixTimeMilliseconds();
            try
            {
                if (_mcpClient == null)
                {
                    EnsureClients(createWordIfMissing: false);
                }

                await EnsureWsBoundAsync(_currentConversationId).ConfigureAwait(true);

                bool isFirstChunk = true;
                string lastToolFinishReason = null;

                var mcpChatResult = await _mcpClient.ChatWithToolsStreamAsync(
                    userInput,
                    null,
                    chunk =>
                    {
                        if (chunk?.choices == null || chunk.choices.Count == 0)
                        {
                            return;
                        }

                        var delta = chunk.choices[0].delta;
                        if (!string.IsNullOrEmpty(delta?.content))
                        {
                            responseBuilder.Append(delta.content);
                        }

                        string reasoningDelta = ExtractReasoningDelta(delta);
                        if (!string.IsNullOrEmpty(reasoningDelta))
                        {
                            reasoningBuilder.Append(reasoningDelta);
                        }

                        var finishReason = chunk.choices[0].finish_reason;
                        if (finishReason == "tool_calls" || finishReason == "function_call")
                        {
                            lastToolFinishReason = finishReason;
                        }

                        object toolCallsDelta = BuildToolCallsDelta(delta);
                        if (!string.IsNullOrEmpty(delta?.content)
                            || toolCallsDelta != null
                            || !string.IsNullOrEmpty(reasoningDelta)
                            || !string.IsNullOrEmpty(finishReason))
                        {
                            Invoke((MethodInvoker)delegate
                            {
                                if (isFirstChunk)
                                {
                                    ResetUploadContext();
                                }

                                _bridge.SendToJavaScript("systemMessage", new
                                {
                                    id = messageId,
                                    content = responseBuilder.ToString(),
                                    reasoningContent = reasoningBuilder.Length > 0
                                        ? reasoningBuilder.ToString()
                                        : null,
                                    toolCallsDelta,
                                    finishReason,
                                    timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                                    isStreaming = true,
                                    isUpdate = !isFirstChunk
                                });
                                isFirstChunk = false;
                            });
                        }
                    },
                    maxIterations: 30,
                    cancellationTokenSource: _cts,
                    conversationId: _currentConversationId,
                    buildHeaders: CreateApiHeaders,
                    resetAuthoritativeUploadContext: ResetUploadContext,
                    onConversationIdKnown: OnConversationIdKnownAsync).ConfigureAwait(true);

                if (mcpChatResult != null
                    && !string.IsNullOrEmpty(mcpChatResult.ConversationId)
                    && mcpChatResult.ConversationId != "-1")
                {
                    _currentConversationId = mcpChatResult.ConversationId;
                    SyncConversationContext();
                    await BindWsAsync(_currentConversationId).ConfigureAwait(true);
                    _bridge.SendToJavaScript("conversationIdChanged", new
                    {
                        conversationId = _currentConversationId
                    });
                }

                string finalContent = responseBuilder.ToString();
                _conversationHistory.Add(new ChatMessage { role = "system", content = finalContent });
                Invoke((MethodInvoker)delegate
                {
                    _bridge.SendToJavaScript("systemMessage", new
                    {
                        id = messageId,
                        content = finalContent,
                        reasoningContent = reasoningBuilder.Length > 0
                            ? reasoningBuilder.ToString()
                            : null,
                        finishReason = lastToolFinishReason,
                        timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                        isStreaming = false,
                        isUpdate = true
                    });
                    CleanupRequestState();
                });
            }
            catch (DailyQuotaExceededException)
            {
                Invoke((MethodInvoker)delegate
                {
                    _bridge.SendToJavaScript("systemMessage", new
                    {
                        id = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                        content = "灵感值已经用完",
                        timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                        isStreaming = false,
                        isHint = true
                    });
                    CleanupRequestState();
                });
            }
            catch (LlmException ex) when (ex.Message.Contains("认证") || ex.Message.Contains("登录"))
            {
                Invoke((MethodInvoker)delegate
                {
                    _bridge.SendToJavaScript("requiresLogin", new
                    {
                        userInput,
                        message = ex.Message
                    });
                    CleanupRequestState();
                });
                _ = LoginForm.ShowDialogAsync(FindForm() ?? (IWin32Window)this, logoutFirst: true);
            }
            catch (OperationCanceledException)
            {
                if (!_newSessionResetPending)
                {
                    string partial = responseBuilder.ToString();
                    if (!string.IsNullOrEmpty(partial))
                    {
                        Invoke((MethodInvoker)delegate
                        {
                            _bridge.SendToJavaScript("systemMessage", new
                            {
                                id = messageId,
                                content = partial,
                                timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                                isStreaming = false,
                                isUpdate = true
                            });
                        });
                    }

                    CleanupRequestState();
                }
            }
            catch (Exception ex)
            {
                Invoke((MethodInvoker)delegate
                {
                    _bridge.SendToJavaScript("systemMessage", new
                    {
                        id = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                        content = "错误: " + ex.Message,
                        timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds(),
                        isStreaming = false
                    });
                    CleanupRequestState();
                });
            }
            finally
            {
                ResetUploadContext();
            }
        }

        private async Task<object> HandleStopRequestAsync(object data)
        {
            try
            {
                if (!string.IsNullOrEmpty(_currentConversationId)
                    && _currentConversationId != "-1"
                    && _wsClient != null)
                {
                    await _wsClient.SendRunCancelAsync(_currentConversationId).ConfigureAwait(false);
                }

                _wsClient?.CancelActiveInvokeLocal();
                if (_cts != null && !_cts.IsCancellationRequested)
                {
                    _cts.Cancel();
                }

                return new { success = true };
            }
            catch (Exception ex)
            {
                return new { success = false, message = ex.Message };
            }
        }

        private async Task<object> HandleAddConversationAsync(object data)
        {
            try
            {
                _newSessionResetPending = true;
                if (_isProcessing && _cts != null && !_cts.IsCancellationRequested)
                {
                    _cts.Cancel();
                }

                _conversationHistory.Clear();
                _currentConversationId = "-1";
                SyncConversationContext();
                _authoritativeUploadId = null;
                if (_wsClient != null)
                {
                    await _wsClient.UnbindAsync().ConfigureAwait(false);
                }

                void Notify()
                {
                    _bridge.SendToJavaScript("clearMessages", new { });
                    CleanupRequestState();
                    _newSessionResetPending = false;
                }

                if (InvokeRequired)
                {
                    Invoke((MethodInvoker)Notify);
                }
                else
                {
                    Notify();
                }

                return new { success = true, conversationId = "-1" };
            }
            catch (Exception ex)
            {
                _newSessionResetPending = false;
                return new { success = false, message = ex.Message };
            }
        }

        private async Task<object> HandleOpenConversationAsync(object data)
        {
            try
            {
                int id = 0;
                if (data is JObject jo)
                {
                    id = jo["id"]?.Value<int>() ?? jo["conversationId"]?.Value<int>() ?? 0;
                }
                else if (data != null)
                {
                    var j = JObject.FromObject(data);
                    id = j["id"]?.Value<int>() ?? j["conversationId"]?.Value<int>() ?? 0;
                }

                if (id <= 0)
                {
                    return new { success = false, message = "无效 conversation id" };
                }

                string baseUrl = ConfigManager.Config.Api.BaseUrl;
                string url = baseUrl.TrimEnd('/') + "/conversations/detail";
                string body = JsonConvert.SerializeObject(new { id });
                var result = await BackendApiClient.PostJsonAuthenticatedAsync(
                    url, body, retryOnUnauthorized: true, timeout: null, authHandleAllErrors: true)
                    .ConfigureAwait(true);
                if (!result.Success)
                {
                    return new { success = false, message = "加载对话失败" };
                }

                var detail = JsonConvert.DeserializeObject<JObject>(result.Body);
                string content = detail?["content"]?.ToString();
                var rawMessages = string.IsNullOrEmpty(content)
                    ? new List<JObject>()
                    : (JsonConvert.DeserializeObject<List<JObject>>(content) ?? new List<JObject>());
                var vueMessages = BuildVueHistoryMessages(rawMessages);

                _currentConversationId = id.ToString();
                SyncConversationContext();
                _conversationHistory.Clear();
                foreach (var m in vueMessages)
                {
                    try
                    {
                        var dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(
                            JsonConvert.SerializeObject(m));
                        if (dict == null)
                        {
                            continue;
                        }

                        string role = dict.ContainsKey("role") ? dict["role"]?.ToString() : null;
                        string text = dict.ContainsKey("content") ? dict["content"]?.ToString() : null;
                        if (!string.IsNullOrEmpty(role))
                        {
                            _conversationHistory.Add(new ChatMessage
                            {
                                role = role,
                                content = text ?? ""
                            });
                        }
                    }
                    catch
                    {
                    }
                }

                _bridge.SendToJavaScript("conversationHistory", new
                {
                    messages = vueMessages,
                    conversationId = _currentConversationId
                });
                await EnsureWsBoundAsync(_currentConversationId).ConfigureAwait(true);
                return new { success = true, conversationId = _currentConversationId };
            }
            catch (Exception ex)
            {
                return new { success = false, message = ex.Message };
            }
        }

        private async Task<object> HandleTodoListReplyAsync(object data)
        {
            try
            {
                var jo = data as JObject ?? (data != null ? JObject.FromObject(data) : null);
                string requestId = jo?["requestId"]?.ToString() ?? jo?["request_id"]?.ToString();
                bool ok = jo?["ok"]?.Value<bool>() ?? false;
                string error = jo?["error"]?.ToString();
                if (string.IsNullOrEmpty(requestId) || _wsClient == null)
                {
                    return new { success = false };
                }

                await _wsClient.SendReplyAsync(requestId, ok, null, error).ConfigureAwait(false);
                return new { success = true };
            }
            catch (Exception ex)
            {
                return new { success = false, error = ex.Message };
            }
        }

        private List<object> BuildVueHistoryMessages(List<JObject> messages)
        {
            return ConversationHistoryMapper.BuildVueHistoryMessages(messages);
        }

        private Dictionary<string, string> CreateApiHeaders(string xConversationId)
        {
            var headers = new Dictionary<string, string>();
            if (!UserService.Instance.CheckLoginStatus())
            {
                return headers;
            }

            var user = UserService.Instance;
            if (!string.IsNullOrEmpty(user.AccessToken))
            {
                headers["Authorization"] = "Bearer " + user.AccessToken;
            }

            if (!string.IsNullOrEmpty(user.UserName))
            {
                headers["X-Username"] = user.UserName;
            }

            headers["X-Conversation-Id"] = !string.IsNullOrEmpty(xConversationId)
                ? xConversationId
                : _currentConversationId;
            if (!string.IsNullOrWhiteSpace(_authoritativeUploadId))
            {
                headers["X-Upload-Id"] = _authoritativeUploadId.Trim();
            }

            return headers;
        }

        private void ResetUploadContext()
        {
            bool had = !string.IsNullOrWhiteSpace(_authoritativeUploadId);
            _authoritativeUploadId = null;
            if (!had || _uploadUiNotified)
            {
                return;
            }

            _uploadUiNotified = true;
            Invoke((MethodInvoker)delegate
            {
                _bridge.SendToJavaScript("uploadAttachmentContextConsumed", new { });
            });
        }

        private void CleanupRequestState()
        {
            AgentRunCancellation.EndRun();
            _isProcessing = false;
            if (_cts != null)
            {
                _cts.Dispose();
                _cts = null;
            }

            _bridge?.SendToJavaScript("requestStateChanged", new { isProcessing = false });
        }

        private async Task OnConversationIdKnownAsync(string conversationId)
        {
            if (string.IsNullOrWhiteSpace(conversationId)
                || conversationId == "-1"
                || conversationId == "0")
            {
                return;
            }

            if (_currentConversationId != conversationId)
            {
                _currentConversationId = conversationId;
                SyncConversationContext();
                _bridge.SendToJavaScript("conversationIdChanged", new { conversationId });
            }

            await EnsureWsBoundAsync(conversationId).ConfigureAwait(false);
        }

        private async Task EnsureWsReadyAsync()
        {
            EnsureClients(createWordIfMissing: false);
            if (_wsClient == null || !UserService.Instance.CheckLoginStatus())
            {
                return;
            }

            var user = UserService.Instance;
            await _wsClient.ConnectAsync(
                ConfigManager.Config.Api.BaseUrl,
                user.AccessToken,
                user.UserName).ConfigureAwait(false);
        }

        private async Task<bool> EnsureWsBoundAsync(string conversationId)
        {
            if (string.IsNullOrWhiteSpace(conversationId)
                || conversationId == "-1"
                || conversationId == "0")
            {
                return false;
            }

            return await BindWsAsync(conversationId).ConfigureAwait(false);
        }

        private async Task<bool> BindWsAsync(string conversationId)
        {
            await EnsureWsReadyAsync().ConfigureAwait(false);
            if (_wsClient == null || !_wsClient.IsConnected)
            {
                return false;
            }

            return await _wsClient.BindAsync(conversationId).ConfigureAwait(false);
        }

        private async Task<object> HandleGetToolAliasAsync(object data)
        {
            try
            {
                string toolName = null;
                if (data is JObject jObj)
                {
                    toolName = jObj["toolName"]?.ToString() ?? jObj["name"]?.ToString();
                }
                else if (data != null)
                {
                    var jo = JObject.FromObject(data);
                    toolName = jo["toolName"]?.ToString() ?? jo["name"]?.ToString();
                }

                if (string.IsNullOrEmpty(toolName))
                {
                    return new { success = false, message = "工具名称不能为空" };
                }

                // 与 Plugin 一致：别名来自 Backend Registry
                await McpToolsInfo.EnsureRegistryAliasesAsync().ConfigureAwait(false);
                string alias = McpToolsInfo.GetToolAlias(toolName);
                return new { success = true, alias };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DesktopChatSurface] 获取工具别名失败: {ex.Message}");
                return new { success = false, message = ex.Message };
            }
        }

        private void OnUserLoggedIn(object sender, UserEventArgs e)
        {
            _ = EnsureWsReadyAsync();
            _ = McpToolsInfo.EnsureRegistryAliasesAsync();
        }

        private void OnUserLoggedOut(object sender, UserEventArgs e)
        {
            _ = HandleAddConversationAsync(null);
        }

        private static object BuildToolCallsDelta(StreamMessage delta)
        {
            if (delta?.tool_calls == null || delta.tool_calls.Count == 0)
            {
                return null;
            }

            var list = new List<Dictionary<string, object>>();
            foreach (var tc in delta.tool_calls)
            {
                var item = new Dictionary<string, object> { { "index", tc.index } };
                if (!string.IsNullOrEmpty(tc.id))
                {
                    item["id"] = tc.id;
                }

                if (tc.function != null)
                {
                    item["function"] = new Dictionary<string, object>
                    {
                        { "name", tc.function.name },
                        { "arguments", tc.function.arguments }
                    };
                }

                list.Add(item);
            }

            return list;
        }

        private static string ExtractReasoningDelta(StreamMessage delta)
        {
            if (delta == null)
            {
                return null;
            }

            var sb = new System.Text.StringBuilder();
            if (!string.IsNullOrEmpty(delta.reasoning_content))
            {
                sb.Append(delta.reasoning_content);
            }

            if (!string.IsNullOrEmpty(delta.reasoning))
            {
                sb.Append(delta.reasoning);
            }

            return sb.Length > 0 ? sb.ToString() : null;
        }
    }
}
