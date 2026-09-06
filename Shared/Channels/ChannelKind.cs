namespace WordAddIn1
{
    /// <summary>
    /// 操作渠道类型。Word / Wps / Excel / Et / Ppt / Wpp / Browser。
    /// </summary>
    public enum ChannelKind
    {
        Word = 0,
        Ppt = 1,
        Excel = 2,
        Wps = 3,
        Et = 4,
        /// <summary>WPS 演示。channel_id 族 p（与 PowerPoint 共用）；与文字 Wps 区分。</summary>
        Wpp = 5,
        /// <summary>浏览器页。channel_id 族 b；track=agent 易写窗 / attach 为 Chrome/Edge。</summary>
        Browser = 6,
    }
}
