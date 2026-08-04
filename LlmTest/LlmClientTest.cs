using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace WordAddIn1
{
    /// <summary>
    /// 大模型 API 客户端类
    /// </summary>
    public class LlmClient
    {
        private const string BaseUrl = "https://api.deepseek.com/v1";
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
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _jsonSerializer = new JavaScriptSerializer();
        }

        /// <summary>
        /// 发送聊天请求
        /// </summary>
        /// <param name="messages">消息列表</param>
        /// <param name="model">模型名称，默认为 deepseek-chat</param>
        /// <param name="temperature">温度参数，控制随机性 (0.0-2.0)</param>
        /// <param name="maxTokens">最大token数</param>
        /// <returns>API响应</returns>
        public async Task<LlmResponse> ChatAsync(
            List<ChatMessage> messages,
            string model = "deepseek-chat",
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

                HttpResponseMessage response = await _httpClient.PostAsync($"{BaseUrl}/conversations/mcp/chat/stream", content);
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
    /// 大模型异常类
    /// </summary>
    public class LlmException : Exception
    {
        public LlmException(string message) : base(message)
        {
        }

        public LlmException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// 大模型客户端使用示例
    /// </summary>
    public static class LlmExample
    {
        /// <summary>
        /// 使用示例 - 异步调用
        /// </summary>
        public static async Task<string> TestChatAsync()
        {
            try
            {
                // 创建客户端实例（使用您提供的API密钥）
                var client = new LlmClient("sk-2e2929af6bde429990eef75c39e6afd7");

                // 发送简单对话请求
                string response = await client.ChatAsync(
                    userMessage: "你好，请介绍一下你自己。",
                    systemMessage: "你是一个专业的AI助手，擅长帮助用户解决问题。"
                );

                return response;
            }
            catch (LlmException ex)
            {
                return $"大模型API调用失败: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"发生未知错误: {ex.Message}";
            }
        }

        /// <summary>
        /// 使用示例 - 多轮对话
        /// </summary>
        public static async Task<string> TestMultiTurnChatAsync()
        {
            try
            {
                var client = new LlmClient("sk-2e2929af6bde429990eef75c39e6afd7");

                // 构建多轮对话消息
                var messages = new List<ChatMessage>
                {
                    new ChatMessage { role = "system", content = "你是一个专业的Word插件助手。" },
                    new ChatMessage { role = "user", content = "如何在Word中插入表格？" },
                    new ChatMessage { role = "assistant", content = "在Word中插入表格，您可以：1. 点击插入选项卡 2. 点击表格按钮 3. 选择表格大小" },
                    new ChatMessage { role = "user", content = "如何设置表格边框？" }
                };

                var response = await client.ChatAsync(messages);
                return response.choices[0].message.content;
            }
            catch (LlmException ex)
            {
                return $"大模型API调用失败: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"发生未知错误: {ex.Message}";
            }
        }
    }
}
