using System;

namespace WordAddIn1
{
    /// <summary>
    /// 当前 MCP 会话 ID，供 WorkspacePathResolver 等使用。
    /// </summary>
    public static class ConversationContext
    {
        private static string _currentId = "-1";

        public static string CurrentId
        {
            get => _currentId;
            set
            {
                string next = string.IsNullOrEmpty(value) ? "-1" : value;
                if (!string.Equals(_currentId, next, StringComparison.Ordinal))
                {
                    WorkspaceReady = false;
                }

                _currentId = next;
            }
        }

        /// <summary>打开对话对账成功后为 true。区内读在未对齐时失败。</summary>
        public static bool WorkspaceReady { get; set; }
    }
}
