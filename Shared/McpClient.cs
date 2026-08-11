using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace WordAddIn1
{
    /// <summary>
    /// MCP client: SSE passthrough to Backend Orchestrator (PR-4, no local [TOOL_CALL] loop).
    /// </summary>
    public class McpClient
    {
        private readonly LlmClient _llmClient;
        private readonly JavaScriptSerializer _jsonSerializer;
        private readonly Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> _toolRegistry;
        private readonly object _wordApplication;

        public class ChatResult
        {
            public string Response { get; set; }
            public int ToolCallCount { get; set; }
        }

        public class McpMessage
        {
            public string Jsonrpc { get; set; } = "2.0";
            public object Id { get; set; }
            public string Method { get; set; }
            public object Params { get; set; }
            public object Result { get; set; }
            public McpError Error { get; set; }
        }

        public class McpError
        {
            public int Code { get; set; }
            public string Message { get; set; }
            public object Data { get; set; }
        }

        public McpClient(string llmApiKey, object wordApplication = null)
        {
            _llmClient = new LlmClient(llmApiKey);
            _jsonSerializer = new JavaScriptSerializer();
            _toolRegistry = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>();
            _wordApplication = wordApplication;
            McpTools.RegisterBuiltInTools(_toolRegistry, _wordApplication);
        }

        public async Task InitializeAsync()
        {
                await RefreshAvailableToolsAsync();
        }

        public async Task RefreshAvailableToolsAsync()
        {
            await Task.CompletedTask;
        }

        public async Task<ChatResult> ChatWithToolsAsync(
            string userMessage,
            string systemMessage = null,
            int maxIterations = 30,
            List<ChatMessage> conversationHistory = null)
        {
            var messages = new List<ChatMessage>();
            if (conversationHistory != null && conversationHistory.Count > 0)
            {
                messages.AddRange(conversationHistory);
            }
            messages.Add(new ChatMessage { role = "user", content = userMessage });

                var response = await _llmClient.ChatAsync(messages);
                if (response.choices == null || response.choices.Count == 0)
                {
                    return new ChatResult
                    {
                    Response = "Sorry, no response was generated.",
                    ToolCallCount = 0
                };
            }

            return new ChatResult
            {
                Response = response.choices[0].message?.content ?? "",
                ToolCallCount = 0
            };
        }

        public async Task<LlmClient.ChatStreamResult> ChatWithToolsStreamAsync(
            string userMessage,
            string systemMessage = null,
            Action<StreamChunk> onChunk = null,
            int maxIterations = 30,
            CancellationTokenSource cancellationTokenSource = null,
            string conversationId = "-1",
            Func<string, Dictionary<string, string>> buildHeaders = null,
            Action resetAuthoritativeUploadContext = null,
            Func<string, Task> onConversationIdKnown = null,
            object openChannels = null)
        {
            var messages = new List<ChatMessage>
            {
                new ChatMessage { role = "user", content = userMessage }
            };

            var headers = buildHeaders?.Invoke(conversationId);
            var token = cancellationTokenSource?.Token ?? default;

            var chatResult = await _llmClient.ChatStreamAsync(
                messages,
                onChunk,
                headers: headers,
                cancellationToken: token,
                onConversationIdKnown: onConversationIdKnown,
                openChannels: openChannels);

            resetAuthoritativeUploadContext?.Invoke();

            return chatResult ?? new LlmClient.ChatStreamResult { ConversationId = conversationId };
        }

        public List<string> GetRegisteredToolNames()
        {
            return _toolRegistry.Keys.ToList();
        }
    }
}
