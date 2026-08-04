namespace WordAddIn1
{
    /// <summary>
    /// 当前进程宿主类型，影响无默认渠道时是否允许回退 ActiveDocument（I6）。
    /// </summary>
    public enum ChannelHostKind
    {
        Unknown = 0,
        Plugin = 1,
        Desktop = 2,
    }

    public static class ChannelHost
    {
        public static ChannelHostKind Kind { get; set; } = ChannelHostKind.Unknown;

        /// <summary>
        /// Plugin：默认可回退 ActiveDocument；Desktop：必须先有渠道。
        /// </summary>
        public static bool AllowActiveDocumentFallback => Kind == ChannelHostKind.Plugin;
    }
}
