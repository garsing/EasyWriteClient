using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// Word 宿主：读内容走现网 <see cref="WordDocumentExtractor"/>。
    /// </summary>
    internal static class WordDocumentHost
    {
        public static GetDocumentContentHostResult GetDocumentContent(
            DocumentSessionContext context,
            string codeLevel)
        {
            if (context == null || context.WordDocument == null)
            {
                throw new ArgumentException("Word 文档上下文无效。");
            }

            Word.Document document = context.WordDocument;
            var processingResult = WordDocumentExtractor.ProcessDocument(
                document,
                ProcessDocumentOptions.ForGetDocumentContent("get_document_content"));

            string displayName = null;
            try
            {
                displayName = document.Name;
            }
            catch (Exception)
            {
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
