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
        /// <summary>WPS 演示（渠道前缀 wpp:；与文字 Wps 区分）。</summary>
        Wpp = 5,
        /// <summary>浏览器页（渠道前缀 browser:agent: / browser:attach:）。</summary>
        Browser = 6,
    }
}
