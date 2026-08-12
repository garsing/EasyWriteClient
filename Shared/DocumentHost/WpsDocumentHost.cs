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

            if (!TryAsWordDocument(context.WpsDocument, out Word.Document wordDoc))
            {
                throw new InvalidOperationException(
                    "unsupported: 当前 WPS 文档无法接入抽取管线（非 Word 兼容自动化对象）；host=wps");
            }

            // 与 WpsDocumentIdentity 对齐 uuid，避免 ProcessDocument 再生成 Word 侧 uuid 分叉
            string wpsUuid = context.DocUuid;
            if (!string.IsNullOrEmpty(wpsUuid))
            {
                DocumentIdentity.Register(wordDoc, wpsUuid);
                DocumentState.GetOrCreateSession(wpsUuid);
                DocumentState.ActivateSessionByUuid(wpsUuid);
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

        private static bool TryAsWordDocument(object doc, out Word.Document wordDoc)
        {
            wordDoc = null;
            if (doc == null)
            {
                return false;
            }

            wordDoc = doc as Word.Document;
            if (wordDoc != null)
            {
                return true;
            }

            try
            {
                wordDoc = (Word.Document)doc;
                return wordDoc != null;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
