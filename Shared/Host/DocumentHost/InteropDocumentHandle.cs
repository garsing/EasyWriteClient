using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// 适配层解析结果：可供现网 Word Interop 管线使用的 Document + 会话元数据。
    /// F_* 用 <see cref="Document"/> 编排业务；用 <see cref="DisallowBackendApi"/> 约束 ProcessDocument，无需写宿主分支。
    /// </summary>
    public sealed class InteropDocumentHandle
    {
        public InteropDocumentHandle(
            Word.Document document,
            DocumentSessionContext context,
            bool disallowBackendApi)
        {
            Document = document;
            Context = context;
            DisallowBackendApi = disallowBackendApi;
        }

        public Word.Document Document { get; }

        public DocumentSessionContext Context { get; }

        /// <summary>WPS 等宿主本期禁止 ProcessDocument 走后端 API 分块（I6b）。</summary>
        public bool DisallowBackendApi { get; }

        public string ChannelId
        {
            get { return ChannelRegistry.ToPublicId(Context?.ChannelId) ?? ""; }
        }

        public string HostName
        {
            get
            {
                return Context == null
                    ? ""
                    : Context.Kind.ToString().ToLowerInvariant();
            }
        }
    }
}
