namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// 适配层读内容成功结果（供 F_get_document_content 组装 ToolResult）。
    /// </summary>
    public sealed class GetDocumentContentHostResult
    {
        public DocumentSessionContext Context { get; set; }

        public WordDocumentExtractor.ProcessingResult ProcessingResult { get; set; }

        public string DocumentDisplayName { get; set; }
    }
}
