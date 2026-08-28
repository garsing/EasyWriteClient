namespace WordAddIn1
{
    public enum ForegroundDanceKind
    {
        None = 0,
        Open = 1,
        Operate = 2,
    }

    /// <summary>
    /// Desktop 前台编排请求：缩 → 目标窗焦点（不 Maximize）→ 易写再焦点。
    /// Plugin 不注册回调，Raise 为 no-op。
    /// </summary>
    public sealed class ForegroundDanceRequest
    {
        public ForegroundDanceKind Kind { get; set; }

        /// <summary>0 = 跳过第 2 步，仍做缩小 + 拉回易写。</summary>
        public int TargetHwnd { get; set; }

        public string ChannelId { get; set; }

        /// <summary>调试用：工具名或调用点。</summary>
        public string Source { get; set; }
    }
}
