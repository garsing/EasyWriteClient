namespace WordAddIn1
{
    /// <summary>
    /// 操作渠道类型。首期仅实现 Word；Ppt/Excel 为预留扩展点（D11/D13）。
    /// </summary>
    public enum ChannelKind
    {
        Word = 0,
        Ppt = 1,
        Excel = 2,
    }
}
