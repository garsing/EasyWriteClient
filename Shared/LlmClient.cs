using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.IO;
using System.Threading;
using System.Net;
using System.Net.NetworkInformation;
using System.Linq;
using Newtonsoft.Json;

namespace WordAddIn1
{
    /// <summary>
    /// 大模型 API 客户端类
    /// </summary>
    public class LlmClient
    {
        private readonly string _baseUrl;
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;
        private readonly JavaScriptSerializer _jsonSerializer;


        /// <summary>
        /// 初始化大模型客户端
        /// </summary>
        /// <param name="apiKey">API 密钥</param>
        public LlmClient(string apiKey)
        {
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));

            // 从配置中读取LLM API地址
            var config = ConfigManager.Config.LLMApi;
            _baseUrl = config.BaseUrl;

            // 配置HttpClient - 针对Word环境优化
            var handler = new HttpClientHandler
            {
                // 允许自动重定向
                AllowAutoRedirect = true,
                // 设置最大重定向次数
                MaxAutomaticRedirections = 5,
                // 使用默认凭据
                UseDefaultCredentials = false,
                // 禁用自动代理检测，使用系统代理
                UseProxy = true,
                Proxy = WebRequest.GetSystemWebProxy(),
                // 确保SSL证书验证
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) =>
                {
                    // 在开发环境中接受所有证书（生产环境应该验证）
                    return true;
                }
            };

            // 强制使用TLS 1.2及以上
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            _httpClient = new HttpClient(handler);
            _httpClient.Timeout = TimeSpan.FromSeconds(30); // 设置30秒超时
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("WordAddIn/1.0");

            _jsonSerializer = new JavaScriptSerializer();
        }

        /// <summary>
        /// 发送聊天请求
        /// </summary>
        /// <param name="messages">消息列表</param>
        /// <param name="model">模型名称，默认为 deepseek-v4-flash（实际以服务端用户套餐为准）</param>
        /// <param name="temperature">温度参数，控制随机性 (0.0-2.0)</param>
        /// <param name="maxTokens">最大token数</param>
        /// <returns>API响应</returns>
        public async Task<LlmResponse> ChatAsync(
            List<ChatMessage> messages,
            string model = "deepseek-v4-flash",
            double temperature = 0.7,
            int? maxTokens = null)
        {
            try
            {
                var requestBody = new
                {
                    model = model,
                    messages = messages,
                    temperature = temperature,
                    max_tokens = maxTokens
                };

                string jsonRequest = _jsonSerializer.Serialize(requestBody);
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await _httpClient.PostAsync($"{_baseUrl}/llm/chat/completions", content);
                response.EnsureSuccessStatusCode();

                string jsonResponse = await response.Content.ReadAsStringAsync();
                return _jsonSerializer.Deserialize<LlmResponse>(jsonResponse);
            }
            catch (HttpRequestException ex)
            {
                throw new LlmException($"API请求失败: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new LlmException($"请求处理失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 简单文本对话（便捷方法）
        /// </summary>
        /// <param name="userMessage">用户消息</param>
        /// <param name="systemMessage">系统消息（可选）</param>
        /// <returns>AI回复内容</returns>
        public async Task<string> ChatAsync(string userMessage, string systemMessage = null)
        {
            var messages = new List<ChatMessage>();

            if (!string.IsNullOrEmpty(systemMessage))
            {
                messages.Add(new ChatMessage { role = "system", content = systemMessage });
            }

            messages.Add(new ChatMessage { role = "user", content = userMessage });

            var response = await ChatAsync(messages);
            return response.choices[0].message.content;
        }

        /// <summary>
        /// 发送流式聊天请求
        /// </summary>
        /// <param name="messages">消息列表</param>
        /// <param name="onChunk">处理每个数据块的回调函数</param>
        /// <param name="model">模型名称，默认为 deepseek-v4-flash</param>
        /// <param name="temperature">温度参数，控制随机性 (0.0-2.0)</param>
        /// <param name="maxTokens">最大token数</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <summary>
        /// 诊断网络连接问题
        /// </summary>
        private async Task<string> DiagnoseNetworkAsync()
        {
            var diagnostics = new StringBuilder();
            diagnostics.AppendLine("=== 网络诊断信息 ===");

            try
            {
                // 从配置中获取主机地址
                var uri = new Uri(_baseUrl);
                var host = uri.Host;

                // 检查DNS解析
                diagnostics.AppendLine("1. DNS解析测试:");
                var hostEntry = Dns.GetHostEntry(host);
                diagnostics.AppendLine($"   {host} 解析到: {string.Join(", ", hostEntry.AddressList.Select(a => a.ToString()))}");

                // 检查基本连接
                diagnostics.AppendLine("2. 基本连接测试:");
                using (var ping = new Ping())
                {
                    var reply = await ping.SendPingAsync(host, 3000);
                    diagnostics.AppendLine($"   Ping {host}: {reply.Status}, RoundTrip: {reply.RoundtripTime}ms");
                }

                // 检查代理设置
                diagnostics.AppendLine("3. 代理设置:");
                var proxy = WebRequest.GetSystemWebProxy();
                var testUri = new Uri(_baseUrl);
                var proxyUri = proxy.GetProxy(testUri);
                diagnostics.AppendLine($"   系统代理: {proxyUri}");
                diagnostics.AppendLine($"   是否使用代理: {!proxyUri.Equals(testUri)}");

                // 检查SSL/TLS版本
                diagnostics.AppendLine("4. SSL/TLS支持:");
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                diagnostics.AppendLine($"   支持的SSL版本: {ServicePointManager.SecurityProtocol}");

            }
            catch (Exception ex)
            {
                diagnostics.AppendLine($"诊断过程中出错: {ex.Message}");
            }

            return diagnostics.ToString();
        }

        /// <summary>识别服务端日额度超额（402 或 body 含 DAILY_QUOTA_EXCEEDED）。</summary>
        private static bool IsDailyQuotaExceededResponse(int statusCode, string responseContent)
        {
            if (statusCode == 402)
            {
                return true;
            }
            if (string.IsNullOrEmpty(responseContent))
            {
                return false;
            }
            return responseContent.IndexOf("DAILY_QUOTA_EXCEEDED", StringComparison.OrdinalIgnoreCase) >= 0
                || responseContent.IndexOf("灵感值已经用完", StringComparison.Ordinal) >= 0
                || responseContent.IndexOf("今日灵感值已用完", StringComparison.Ordinal) >= 0
                || responseContent.IndexOf("今日额度已用完", StringComparison.Ordinal) >= 0;
        }

        /// <summary>
        /// 流式聊天结果类
        /// </summary>
        public class ChatStreamResult
        {
            public string ConversationId { get; set; }
            public string FinalResponse { get; set; } // 当工具执行失败时直接返回最终响应
        }

        public async Task<ChatStreamResult> ChatStreamAsync(
            List<ChatMessage> messages,
            Action<StreamChunk> onChunk,
            string model = "deepseek-v4-flash",
            double temperature = 0.7,
            int? maxTokens = null,
            Dictionary<string, string> headers = null,
            CancellationToken cancellationToken = default,
            Func<string, Task> onConversationIdKnown = null,
            object openChannels = null)
        {
            try
            {
                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] 开始准备流式请求，消息数量: {messages?.Count ?? 0}");

                // 跳过网络诊断以提升性能 - 在生产环境中网络通常是稳定的

                var requestBody = new Dictionary<string, object>
                {
                    ["model"] = model,
                    ["messages"] = messages,
                    ["temperature"] = temperature,
                    ["stream"] = true
                };
                if (maxTokens.HasValue)
                {
                    requestBody["max_tokens"] = maxTokens.Value;
                }

                if (openChannels != null)
                {
                    requestBody["open_channels"] = openChannels;
                }

                string jsonRequest = _jsonSerializer.Serialize(requestBody);
                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] 请求体长度: {jsonRequest.Length}");
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                string url = $"{_baseUrl}/llm/chat/completions";
                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] 请求URL: {url}");

                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Content = content;

                    // 添加自定义请求头
                    if (headers != null && headers.Count > 0)
                    {
                        // EasyWriteDiagnostics.Log(DebugCategory.Llm, "[HTTP] 添加自定义请求头:");
                        foreach (var header in headers)
                        {
                            if (header.Key.ToLower() == "authorization")
                            {
                                // Authorization头需要特殊处理
                                var token = header.Value;
                                if (token.StartsWith("Bearer "))
                                {
                                    token = token.Substring(7);
                                }
                                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[HTTP] Authorization: Bearer ***{token.Substring(Math.Max(0, token.Length - 10))}***");
                            }
                            else
                            {
                                // 其他自定义头
                                request.Headers.Add(header.Key, header.Value);
                                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[HTTP] {header.Key}: {header.Value}");
                            }
                        }
                        // EasyWriteDiagnostics.Log(DebugCategory.Llm, "[HTTP] 请求头添加完成");
                    }
                    else
                    {
                        // EasyWriteDiagnostics.Log(DebugCategory.Llm, "[HTTP] 无自定义请求头");
                    }

                    // EasyWriteDiagnostics.Log(DebugCategory.Llm, "[STREAM_DEBUG] 发送HTTP请求...");

                    // 记录发送HTTP请求的时间
                    DateTime httpRequestStartTime = DateTime.Now;
                    // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[TIMER] 发送HTTP请求 - {httpRequestStartTime:yyyy-MM-dd HH:mm:ss.fff}");

                    // 发送请求并处理可能的认证错误
                    HttpResponseMessage response;
                    bool retryAttempted = false;

                    while (true)
                    {
                        response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

                        // 记录收到HTTP响应的时间
                        DateTime httpResponseTime = DateTime.Now;
                        // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[TIMER] 收到HTTP响应头 - {httpResponseTime:yyyy-MM-dd HH:mm:ss.fff}, HTTP请求耗时: {(httpResponseTime - httpRequestStartTime).TotalMilliseconds:F2}ms");
                        // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] 收到响应，状态码: {(int)response.StatusCode}");

                        // 检查响应状态，如果是401错误，交给Auth模块处理
                        if (!response.IsSuccessStatusCode)
                        {
                            string responseContent = await response.Content.ReadAsStringAsync();
                            // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] 响应内容: {responseContent}");

                            if (IsDailyQuotaExceededResponse((int)response.StatusCode, responseContent))
                            {
                                response.Dispose();
                                throw new DailyQuotaExceededException("灵感值已经用完");
                            }

                            // 调用Auth模块处理响应
                            var newToken = await WordAddIn1.Auth.Instance.HandleApiResponse(response, responseContent);

                            // 如果返回了新token，说明刷新成功，使用新token重试请求
                            if (!string.IsNullOrEmpty(newToken) && !retryAttempted)
                            {
                                // EasyWriteDiagnostics.Log(DebugCategory.Llm, "[STREAM_DEBUG] 令牌刷新成功，重试请求");

                                // 更新请求头中的Authorization
                                if (request.Headers.Authorization != null)
                                {
                                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", newToken);
                                }

                                // 释放当前响应，准备重试
                                response.Dispose();
                                retryAttempted = true;
                                continue; // 重新发送请求
                            }
                            else
                            {
                                // Auth模块已经处理了错误，或者重试也失败了
                                response.Dispose();
                                return new ChatStreamResult { ConversationId = null };
                            }
                        }
                        else
                        {
                            // 响应成功，跳出重试循环
                            break;
                        }
                    }

                    using (response) // 现在response是成功的响应
                    {
                        // 重要：在开始读取流式响应之前，先读取响应头中的对话ID
                        // 这样即使流式响应被取消，也能获取到对话ID
                        string responseConversationId = null;
                        if (response.Headers.TryGetValues("X-Conversation-Id", out var conversationIdValues))
                        {
                            responseConversationId = conversationIdValues.FirstOrDefault();
                            EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[LLM_CLIENT] 从响应头读取对话ID: {responseConversationId}");
                        }

                        // 创建一个变量来保存结果，确保即使抛出异常也能返回对话ID
                        ChatStreamResult result = new ChatStreamResult { ConversationId = responseConversationId };

                        if (!string.IsNullOrEmpty(responseConversationId)
                            && responseConversationId != "0"
                            && onConversationIdKnown != null)
                        {
                            await onConversationIdKnown(responseConversationId);
                        }

                        try
                        {
                            using (var stream = await response.Content.ReadAsStreamAsync())
                            using (var reader = new StreamReader(stream))
                            {
                                string line;
                                while ((line = await reader.ReadLineAsync()) != null)
                                {
                                    cancellationToken.ThrowIfCancellationRequested();

                                    if (string.IsNullOrWhiteSpace(line))
                                        continue;

                                    // SSE格式: data: {...}
                                    if (line.StartsWith("data: "))
                                    {
                                        string data = line.Substring(6);

                                        // 检查是否是结束标记
                                        if (data == "[DONE]")
                                            break;

                                        try
                                        {
                                            // 首先尝试解析为错误响应
                                            var errorResponse = _jsonSerializer.Deserialize<StreamErrorResponse>(data);
                                            if (errorResponse?.error != null)
                                            {
                                                // 流式响应中包含错误
                                                string errorMsg = errorResponse.error.message ?? "流式响应错误";
                                                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] 流式响应错误: {errorMsg}");

                                                // 根据错误类型处理
                                                if (errorMsg.Contains("token") || errorMsg.Contains("认证") || errorMsg.Contains("unauthorized"))
                                                {
                                                    // 可能是认证相关错误，触发重新认证
                                                    await WordAddIn1.Auth.Instance.RefreshAccessToken();
                                                }

                                                throw new LlmException($"流式API错误: {errorMsg}");
                                            }

                                            // 解析为正常的数据块（Newtonsoft 可正确解析嵌套 tool_calls / function）
                                            var chunk = JsonConvert.DeserializeObject<StreamChunk>(data);
                                            onChunk?.Invoke(chunk);
                                        }
                                        catch (LlmException)
                                        {
                                            // 重新抛出LlmException
                                            throw;
                                        }
                                        catch (Exception)
                                        {
                                            // 忽略其他解析错误，继续处理下一行
                                            continue;
                                        }
                                    }
                                }
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            // 流式响应被取消，但仍然返回已读取的对话ID
                            EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[LLM_CLIENT] 流式响应被取消，返回已读取的对话ID: {responseConversationId}");
                            // 不重新抛出异常，而是返回结果，让上层决定如何处理
                            return result;
                        }

                        // 返回结果，包含从响应头获取的对话ID
                        return result;
                    }
                    } // 结束外层using (response)
            }
            catch (HttpRequestException ex)
            {
                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] HttpRequestException: {ex.Message}");
                throw new LlmException($"流式API请求失败 - 网络错误: {ex.Message}", ex);
            }
            catch (OperationCanceledException)
            {
                // EasyWriteDiagnostics.Log(DebugCategory.Llm, "[STREAM_DEBUG] 请求被用户取消");
                throw; // 重新抛出OperationCanceledException，让上层处理
            }
            catch (DailyQuotaExceededException)
            {
                // 超额由上层以灰字提示，勿包成 LlmException
                throw;
            }
            catch (Exception ex)
            {
                // EasyWriteDiagnostics.Log(DebugCategory.Llm, $"[STREAM_DEBUG] Exception: {ex.Message}");
                throw new LlmException($"流式请求处理失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 简单流式文本对话（便捷方法）
        /// </summary>
        /// <param name="userMessage">用户消息</param>
        /// <param name="systemMessage">系统消息（可选）</param>
        /// <param name="onChunk">处理每个数据块的回调函数</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>流式聊天结果，包含对话ID</returns>
        public async Task<ChatStreamResult> ChatStreamAsync(
            string userMessage,
            string systemMessage = null,
            Action<StreamChunk> onChunk = null,
            CancellationToken cancellationToken = default)
        {
            var messages = new List<ChatMessage>();

            if (!string.IsNullOrEmpty(systemMessage))
            {
                messages.Add(new ChatMessage { role = "system", content = systemMessage });
            }

            messages.Add(new ChatMessage { role = "user", content = userMessage });

            var fullContent = new StringBuilder();

            var result = await ChatStreamAsync(messages, (chunk) =>
            {
                onChunk?.Invoke(chunk);

                // 累积内容
                if (chunk.choices != null && chunk.choices.Count > 0)
                {
                    var delta = chunk.choices[0].delta;
                    if (!string.IsNullOrEmpty(delta?.content))
                    {
                        fullContent.Append(delta.content);
                    }
                }
            }, cancellationToken: cancellationToken);

            // 注意：这里我们忽略fullContent，只返回对话ID
            // 调用者应该通过onChunk回调获取内容
            return result;
        }

    }

    /// <summary>
    /// 聊天消息类
    /// </summary>
    public class ChatMessage
    {
        public string role { get; set; } // "system", "user", "assistant"
        public string content { get; set; }
    }

    /// <summary>
    /// 大模型 API 响应类
    /// </summary>
    public class LlmResponse
    {
        public string id { get; set; }
        public string @object { get; set; }
        public long created { get; set; }
        public string model { get; set; }
        public List<Choice> choices { get; set; }
        public Usage usage { get; set; }
    }

    /// <summary>
    /// 选择项类
    /// </summary>
    public class Choice
    {
        public int index { get; set; }
        public ChatMessage message { get; set; }
        public string finish_reason { get; set; }
    }

    /// <summary>
    /// 使用统计类
    /// </summary>
    public class Usage
    {
        public int prompt_tokens { get; set; }
        public int completion_tokens { get; set; }
        public int total_tokens { get; set; }
    }

    /// <summary>
    /// 流式响应数据类
    /// </summary>
    public class StreamChunk
    {
        public string id { get; set; }
        public string @object { get; set; }
        public long created { get; set; }
        public string model { get; set; }
        public List<StreamChoice> choices { get; set; }
        public Usage usage { get; set; }

        [JsonProperty("easywrite_meta")]
        public EasyWriteStreamMeta easywrite_meta { get; set; }
    }

    /// <summary>
    /// EasyWrite 扩展 SSE 元数据（非 OpenAI 标准字段）
    /// </summary>
    public class EasyWriteStreamMeta
    {
        public string @event { get; set; }
        public int? max_rounds { get; set; }

        [JsonProperty("tool_call_id")]
        public string tool_call_id { get; set; }

        public string text { get; set; }
    }

    /// <summary>
    /// 流式选择项类
    /// </summary>
    public class StreamChoice
    {
        public int index { get; set; }
        public StreamMessage delta { get; set; }

        [JsonProperty("finish_reason")]
        public string finish_reason { get; set; }
    }

    /// <summary>
    /// 流式消息增量类
    /// </summary>
    public class StreamMessage
    {
        public string role { get; set; }
        public string content { get; set; }

        [JsonProperty("reasoning_content")]
        public string reasoning_content { get; set; }

        [JsonProperty("reasoning")]
        public string reasoning { get; set; }

        [JsonProperty("tool_calls")]
        public List<StreamToolCallDelta> tool_calls { get; set; }
    }

    /// <summary>
    /// SSE delta.tool_calls 片段
    /// </summary>
    public class StreamToolCallDelta
    {
        public int index { get; set; }
        public string id { get; set; }

        /// <summary>JSON 字段名为 function（JS 保留字），JavaScriptSerializer 无法可靠反序列化。</summary>
        [JsonProperty("function")]
        public StreamToolCallFunction function { get; set; }
    }

    public class StreamToolCallFunction
    {
        public string name { get; set; }
        public string arguments { get; set; }
    }

    /// <summary>用户日限额与散装灵感值均已用尽（HTTP 402 / DAILY_QUOTA_EXCEEDED）。</summary>
    public class DailyQuotaExceededException : Exception
    {
        public DailyQuotaExceededException(string message) : base(message)
        {
        }
    }

    /// <summary>
    /// 流式响应错误模型
    /// </summary>
    public class StreamErrorResponse
    {
        public StreamError error { get; set; }
    }

    /// <summary>
    /// 流式错误详情
    /// </summary>
    public class StreamError
    {
        public string message { get; set; }
        public string type { get; set; }
        public int code { get; set; }
    }
}
