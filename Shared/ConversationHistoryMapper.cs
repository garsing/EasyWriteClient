using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 将后端会话 content（OpenAI 风格 messages）映射为 Vue 聊天气泡。
    /// assistant.tool_calls → segments.type=toolCall（与实时 ToolCallAccumulator 一致）。
    /// </summary>
    public static class ConversationHistoryMapper
    {
        public static List<object> BuildVueHistoryMessages(IEnumerable<JObject> messages)
        {
            var result = new List<object>();
            if (messages == null)
            {
                return result;
            }

            var messageList = new List<JObject>();
            foreach (JObject message in messages)
            {
                if (message != null)
                {
                    messageList.Add(message);
                }
            }

            Dictionary<string, string> toolOutputs = IndexToolProcessOutputs(messageList);

            long idBase = DateTimeOffset.Now.ToUnixTimeMilliseconds();
            int seq = 0;
            int toolSeq = 0;

            foreach (JObject message in messageList)
            {
                if (message == null)
                {
                    continue;
                }

                string role = message["role"]?.ToString();
                if (string.IsNullOrEmpty(role) || role == "system" || role == "tool")
                {
                    continue;
                }

                string text = ContentToDisplayString(message["content"]);
                string reasoning = message["reasoning_content"]?.ToString()
                    ?? message["reasoning"]?.ToString();

                if (role == "user")
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    result.Add(new
                    {
                        id = idBase + seq++,
                        role = "user",
                        content = text,
                        timestamp = idBase + seq,
                        isStreaming = false
                    });
                    continue;
                }

                if (role != "assistant")
                {
                    continue;
                }

                var segments = new List<object>();
                if (!string.IsNullOrEmpty(reasoning))
                {
                    segments.Add(new
                    {
                        type = "thinking",
                        content = reasoning,
                        isComplete = true
                    });
                }

                if (!string.IsNullOrEmpty(text))
                {
                    segments.Add(new
                    {
                        type = "text",
                        content = text
                    });
                }

                AppendToolCallSegments(message["tool_calls"] as JArray, segments, ref toolSeq, toolOutputs);

                if (segments.Count == 0)
                {
                    continue;
                }

                result.Add(new
                {
                    id = idBase + seq++,
                    role = "system",
                    content = text ?? string.Empty,
                    segments,
                    timestamp = idBase + seq,
                    isStreaming = false
                });
            }

            return result;
        }

        private static void AppendToolCallSegments(
            JArray toolCalls,
            List<object> segments,
            ref int toolSeq,
            Dictionary<string, string> toolOutputs)
        {
            if (toolCalls == null || toolCalls.Count == 0)
            {
                return;
            }

            foreach (JToken tc in toolCalls)
            {
                if (tc == null || tc.Type == JTokenType.Null)
                {
                    continue;
                }

                JToken function = tc["function"];
                string name = function?["name"]?.ToString()
                    ?? tc["name"]?.ToString()
                    ?? string.Empty;
                string args = ExtractToolArguments(function?["arguments"] ?? tc["arguments"]);
                string toolCallId = tc["id"]?.ToString() ?? string.Empty;
                string outputText = "";
                if (!string.IsNullOrEmpty(toolCallId)
                    && toolOutputs != null
                    && toolOutputs.TryGetValue(toolCallId, out string stored)
                    && !string.IsNullOrEmpty(stored))
                {
                    outputText = stored;
                }

                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(args))
                {
                    continue;
                }

                segments.Add(new
                {
                    type = "toolCall",
                    toolName = name,
                    toolCallId,
                    content = args ?? string.Empty,
                    outputText,
                    isComplete = true,
                    startIndex = toolSeq++
                });
            }
        }

        /// <summary>
        /// 回放只抽 stdout/stderr；交错顺序回不来，stdout 接 stderr、不加分隔线。
        /// </summary>
        private static Dictionary<string, string> IndexToolProcessOutputs(IList<JObject> messages)
        {
            var map = new Dictionary<string, string>();
            if (messages == null)
            {
                return map;
            }

            foreach (JObject message in messages)
            {
                if (message == null || message["role"]?.ToString() != "tool")
                {
                    continue;
                }

                string id = message["tool_call_id"]?.ToString();
                if (string.IsNullOrEmpty(id))
                {
                    continue;
                }

                string text = ExtractProcessOutput(message["content"]);
                if (!string.IsNullOrEmpty(text))
                {
                    map[id] = text;
                }
            }

            return map;
        }

        private static string ExtractProcessOutput(JToken content)
        {
            if (content == null || content.Type == JTokenType.Null)
            {
                return "";
            }

            string raw = content.Type == JTokenType.String
                ? content.ToString()
                : content.ToString(Formatting.None);
            if (string.IsNullOrWhiteSpace(raw))
            {
                return "";
            }

            try
            {
                JToken parsed = JToken.Parse(raw);
                if (parsed is JObject obj)
                {
                    return JoinStdoutStderr(obj);
                }
            }
            catch (JsonException)
            {
            }

            return "";
        }

        private static string JoinStdoutStderr(JObject obj)
        {
            if (obj == null)
            {
                return "";
            }

            string stdout = TokenString(obj["stdout"]);
            string stderr = TokenString(obj["stderr"]);
            if (string.IsNullOrEmpty(stdout) && string.IsNullOrEmpty(stderr) && obj["data"] is JObject data)
            {
                stdout = TokenString(data["stdout"]);
                stderr = TokenString(data["stderr"]);
            }

            if (string.IsNullOrEmpty(stdout))
            {
                return stderr ?? "";
            }

            if (string.IsNullOrEmpty(stderr))
            {
                return stdout;
            }

            return stdout + stderr;
        }

        private static string TokenString(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return "";
            }

            return token.Type == JTokenType.String ? token.ToString() : token.ToString(Formatting.None);
        }

        private static string ExtractToolArguments(JToken argumentsToken)
        {
            if (argumentsToken == null || argumentsToken.Type == JTokenType.Null)
            {
                return string.Empty;
            }

            if (argumentsToken.Type == JTokenType.String)
            {
                return argumentsToken.ToString();
            }

            return argumentsToken.ToString(Formatting.None);
        }

        private static string ContentToDisplayString(JToken content)
        {
            if (content == null || content.Type == JTokenType.Null)
            {
                return null;
            }

            if (content.Type == JTokenType.String)
            {
                return content.ToString();
            }

            if (content.Type == JTokenType.Array)
            {
                var parts = new List<string>();
                foreach (JToken part in (JArray)content)
                {
                    if (part == null)
                    {
                        continue;
                    }

                    string type = part["type"]?.ToString();
                    if (type == "text" || string.IsNullOrEmpty(type))
                    {
                        string t = part["text"]?.ToString();
                        if (!string.IsNullOrEmpty(t))
                        {
                            parts.Add(t);
                        }
                    }
                }

                return parts.Count > 0 ? string.Join("", parts) : content.ToString(Formatting.None);
            }

            return content.ToString(Formatting.None);
        }
    }
}
