namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 操作前所在文档层。必须在清表 / Replace 之前从 ref 抄下来。
    /// </summary>
    internal sealed class BrowserResnapshotLayer
    {
        public string FrameId { get; set; }

        public string CdpSessionId { get; set; }

        public int? AttachFrameId { get; set; }

        public static BrowserResnapshotLayer FromEntry(BrowserRefEntry entry)
        {
            if (entry == null)
            {
                return new BrowserResnapshotLayer();
            }

            return new BrowserResnapshotLayer
            {
                FrameId = entry.FrameId,
                CdpSessionId = entry.CdpSessionId,
                AttachFrameId = entry.AttachFrameId
            };
        }

        public static BrowserResnapshotLayer Shell()
        {
            return new BrowserResnapshotLayer();
        }
    }
}
