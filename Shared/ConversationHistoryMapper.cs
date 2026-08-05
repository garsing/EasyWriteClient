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

            long idBase = DateTimeOffset.Now.ToUnixTimeMilliseconds();
            int seq = 0;
            int toolSeq = 0;

            foreach (JObject message in messages)
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

                AppendToolCallSegments(message["tool_calls"] as JArray, segments, ref toolSeq);

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
            ref int toolSeq)
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

                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(args))
                {
                    continue;
                }

                segments.Add(new
                {
                    type = "toolCall",
                    toolName = name,
                    content = args ?? string.Empty,
                    isComplete = true,
                    startIndex = toolSeq++
                });
            }
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
