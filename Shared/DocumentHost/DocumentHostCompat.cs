using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// WPS → Word 兼容 RCW 辅助（渠道仍为 wps:，不登记 word:）。
    /// </summary>
    internal static class DocumentHostCompat
    {
        public static bool TryAsWordDocument(object doc, out Word.Document wordDoc)
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

        /// <summary>
        /// 将 WPS 文档对齐为可跑现网 Word Interop 管线的 Document，并同步 uuid/session。
        /// </summary>
        public static bool TryPrepareWpsInteropDocument(
            DocumentSessionContext context,
            out Word.Document wordDoc,
            out string error)
        {
            wordDoc = null;
            error = null;

            if (context == null || context.WpsDocument == null)
            {
                error = "unsupported: WPS 文档上下文无效；host=wps";
                return false;
            }

            if (!TryAsWordDocument(context.WpsDocument, out wordDoc))
            {
                error = "unsupported: 当前 WPS 文档无法接入 Interop 管线（非 Word 兼容自动化对象）；host=wps";
                return false;
            }

            string wpsUuid = context.DocUuid;
            if (!string.IsNullOrEmpty(wpsUuid))
            {
                DocumentIdentity.Register(wordDoc, wpsUuid);
                DocumentState.GetOrCreateSession(wpsUuid);
                DocumentState.ActivateSessionByUuid(wpsUuid);
            }

            try
            {
                wordDoc.Activate();
            }
            catch (Exception)
            {
            }

            return true;
        }
    }
}
