using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// WPS 宿主：晚绑定文档；优先兼容复用 Word 抽取管线（渠道仍为 wps:）。
    /// 后端 API 分块本期禁止（I6b）。
    /// </summary>
    internal static class WpsDocumentHost
    {
        public static GetDocumentContentHostResult GetDocumentContent(
            DocumentSessionContext context,
            string codeLevel)
        {
            if (context == null || context.WpsDocument == null)
            {
                throw new ArgumentException("WPS 文档上下文无效。");
            }

            // 只读：不 Activate，避免读内容时把 WPS 抢到前台
            if (!DocumentHostCompat.TryPrepareWpsInteropDocument(
                    context,
                    out Word.Document wordDoc,
                    out string prepareError,
                    activateDocument: false))
            {
                throw new InvalidOperationException(
                    string.IsNullOrEmpty(prepareError)
                        ? "unsupported: 当前 WPS 文档无法接入抽取管线；host=wps"
                        : prepareError);
            }

            var options = ProcessDocumentOptions.ForGetDocumentContent("get_document_content");
            options.DisallowBackendApi = true;

            WordDocumentExtractor.ProcessingResult processingResult;
            try
            {
                processingResult = WordDocumentExtractor.ProcessDocument(wordDoc, options);
            }
            catch (InvalidOperationException ex)
            {
                if (ex.Message != null
                    && ex.Message.IndexOf("unsupported", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    throw new InvalidOperationException(
                        "unsupported: WPS 大文档后端 API 分块本期未接入（后续再处理）；host=wps。详情: "
                        + ex.Message);
                }

                throw;
            }

            string displayName = null;
            try
            {
                displayName = wordDoc.Name;
            }
            catch (Exception)
            {
                displayName = OpenFiles.WpsCom.TryReadName(context.WpsDocument);
            }

            return new GetDocumentContentHostResult
            {
                Context = context,
                ProcessingResult = processingResult,
                DocumentDisplayName = displayName
            };
        }
    }
}
