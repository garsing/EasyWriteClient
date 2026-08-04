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
            set => _currentId = string.IsNullOrEmpty(value) ? "-1" : value;
        }
    }
}
