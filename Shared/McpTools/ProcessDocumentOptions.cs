namespace WordAddIn1
{
    /// <summary>
    /// ProcessDocument 格式抽取模式（get=None 自 2026-08；粗迁=Full；process_actions=None）。
    /// BodyOnly 枚举值保留；get 入口已停用（见 ForGetDocumentContent）。
    /// </summary>
    public enum ExtractFormatMode
    {
        None = 0,
        BodyOnly = 1,
        Full = 2
    }

    /// <summary>
    /// WordDocumentExtractor.ProcessDocument 行为开关。
    /// </summary>
    public sealed class ProcessDocumentOptions
    {
        public ExtractFormatMode ExtractFormatMode { get; set; } = ExtractFormatMode.Full;

        /// <summary>
        /// 兼容旧调用：true→Full，false→None。新代码请用 <see cref="ExtractFormatMode"/>。
        /// </summary>
        public bool ExtractFormat
        {
            get { return ExtractFormatMode == ExtractFormatMode.Full; }
            set { ExtractFormatMode = value ? ExtractFormatMode.Full : ExtractFormatMode.None; }
        }

        public string SnapshotSource { get; set; }

        /// <summary>为 false 时不 SaveAs/Copy、不抽图落盘，仍插 &lt;image&gt; 占位以保持 seg 结构。</summary>
        public bool ExtractImages { get; set; } = true;

        /// <summary>为 false 时不打印全文/seg/ChunkInfo 等全量 Debug。</summary>
        public bool VerboseDebug { get; set; } = false;

        /// <summary>为 false 时跳过第七步 display_content（process_actions 不需要）。</summary>
        public bool BuildDisplayContent { get; set; } = true;

        /// <summary>为 true 时跳过快照缓存，强制全量读盘（仅代码/调试，MCP 不暴露）。</summary>
        public bool ForceRefresh { get; set; } = false;

        /// <summary>
        /// 为 true 时禁止走后端 API 分块（WPS 本期，见文档操作适配层 I6b）。
        /// 超阈值时 <see cref="WordDocumentExtractor.ProcessDocument"/> 抛 <see cref="System.InvalidOperationException"/>。
        /// </summary>
        public bool DisallowBackendApi { get; set; } = false;

        public static ProcessDocumentOptions ForProcessActions(string snapshotSource)
        {
            return new ProcessDocumentOptions
            {
                ExtractFormatMode = ExtractFormatMode.None,
                SnapshotSource = snapshotSource,
                ExtractImages = false,
                VerboseDebug = false,
                BuildDisplayContent = true
            };
        }

        public static ProcessDocumentOptions ForGetDocumentContent(string snapshotSource)
        {
            // 2026-08：改字显式格式后 get 不再抽 BodyFormat（BodyOnly 代码保留在 WordDocumentExtractor）
            return new ProcessDocumentOptions
            {
                ExtractFormatMode = ExtractFormatMode.None,
                SnapshotSource = snapshotSource,
                ExtractImages = false,
                VerboseDebug = false,
                BuildDisplayContent = true
            };
        }

        /// <summary>粗迁/格式转移：display→pool→全量 format，不抽图、轻量日志。</summary>
        public static ProcessDocumentOptions ForFormatTransfer(string snapshotSource = "format_transfer")
        {
            return new ProcessDocumentOptions
            {
                ExtractFormatMode = ExtractFormatMode.Full,
                SnapshotSource = snapshotSource,
                ExtractImages = false,
                VerboseDebug = false,
                BuildDisplayContent = true
            };
        }
    }
}
