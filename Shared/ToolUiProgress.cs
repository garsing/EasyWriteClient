using System;

namespace WordAddIn1
{
    /// <summary>
    /// 工具执行进度：C# 执行器推到 Vue 折叠标题（如 1/7）。
    /// </summary>
    public static class ToolUiProgress
    {
        public static Action<string, object> SendToJavaScript { get; set; }

        public static void Report(string toolCallId, int current, int total)
        {
            if (string.IsNullOrEmpty(toolCallId) || total <= 1)
            {
                return;
            }

            Action<string, object> send = SendToJavaScript;
            if (send == null)
            {
                return;
            }

            int cur = current;
            if (cur < 1)
            {
                cur = 1;
            }

            if (cur > total)
            {
                cur = total;
            }

            try
            {
                send("toolProgress", new
                {
                    tool_call_id = toolCallId,
                    current = cur,
                    total
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ToolUiProgress] " + ex.Message);
            }
        }
    }
}
