using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// 宿主无关文档会话上下文（由渠道解析得到）。
    /// </summary>
    public sealed class DocumentSessionContext
    {
        public DocumentSessionContext(
            string channelId,
            ChannelKind kind,
            string docUuid,
            Word.Document wordDocument,
            object wpsDocument)
        {
            ChannelId = channelId ?? "";
            Kind = kind;
            DocUuid = docUuid ?? "";
            WordDocument = wordDocument;
            WpsDocument = wpsDocument;
        }

        public string ChannelId { get; }

        public ChannelKind Kind { get; }

        public string DocUuid { get; }

        public Word.Document WordDocument { get; }

        public object WpsDocument { get; }
    }
}
