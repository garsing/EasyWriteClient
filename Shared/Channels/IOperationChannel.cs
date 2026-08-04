namespace WordAddIn1
{
    /// <summary>
    /// 操作渠道：工具解析「操作谁」的显式句柄（D11）。
    /// </summary>
    public interface IOperationChannel
    {
        string ChannelId { get; }
        ChannelKind Kind { get; }
    }
}
