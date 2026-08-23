using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// I4 / I11 / I12：在源分片上按需抽四类产物。指纹变了先清空四类。
    /// </summary>
    public static class FormatSourceEnsure
    {
        public static void RunOnSource(Word.Document sourceDoc, Word.Document targetDoc, Action action)
        {
            try
            {
                if (sourceDoc != null)
                {
                    DocumentState.BindAndActivate(sourceDoc);
                }

                action();
            }
            finally
            {
                if (targetDoc != null)
                {
                    DocumentState.BindAndActivate(targetDoc);
                }
            }
        }

        public static void AlignFingerprint(Word.Document sourceDoc, DocumentSessionState session)
        {
            if (sourceDoc == null || session == null)
            {
                return;
            }

            string fp = BodyXmlFingerprint.Compute(sourceDoc);
            if (string.IsNullOrEmpty(session.FormatExtractBodyFingerprint)
                || !string.Equals(session.FormatExtractBodyFingerprint, fp, StringComparison.Ordinal))
            {
                session.ClearFormatExtractProducts();
                session.FormatExtractBodyFingerprint = fp;
            }
        }

        public static void EnsureCharMapping(Word.Document sourceDoc, bool forceRefresh, bool disallowBackendApi = false)
        {
            if (sourceDoc == null)
            {
                return;
            }

            DocumentSessionState session = DocumentState.GetOrCreateSession(DocumentIdentity.EnsureUuid(sourceDoc));
            AlignFingerprint(sourceDoc, session);
            if (!forceRefresh && session.SubtypeFormatMapping != null && session.SubtypeFormatMapping.Count > 0)
            {
                return;
            }

            ProcessDocumentOptions options = ProcessDocumentOptions.ForFormatTransfer("source_char_mapping");
            options.DisallowBackendApi = disallowBackendApi;
            WordDocumentExtractor.ProcessDocument(sourceDoc, options);
        }

        public static void EnsurePageLayout(Word.Document sourceDoc, bool forceRefresh)
        {
            if (sourceDoc == null)
            {
                return;
            }

            DocumentSessionState session = DocumentState.GetOrCreateSession(DocumentIdentity.EnsureUuid(sourceDoc));
            AlignFingerprint(sourceDoc, session);
            if (!forceRefresh && session.SectionPageLayout != null && session.SectionPageLayout.Count > 0)
            {
                return;
            }

            session.SectionPageLayout = PageLayoutExtractor.Extract(sourceDoc);
        }

        public static void EnsureTableCatalog(Word.Document sourceDoc, bool forceRefresh, bool disallowBackendApi = false)
        {
            if (sourceDoc == null)
            {
                return;
            }

            DocumentSessionState session = DocumentState.GetOrCreateSession(DocumentIdentity.EnsureUuid(sourceDoc));
            AlignFingerprint(sourceDoc, session);
            if (!forceRefresh && session.TableFormatCatalog != null)
            {
                return;
            }

            SourceTableFormatCatalogExtractor.ExtractIntoSession(sourceDoc, session, disallowBackendApi);
        }

        public static void EnsureParaClusters(Word.Document sourceDoc, bool forceRefresh, bool disallowBackendApi = false)
        {
            if (sourceDoc == null)
            {
                return;
            }

            DocumentSessionState session = DocumentState.GetOrCreateSession(DocumentIdentity.EnsureUuid(sourceDoc));
            AlignFingerprint(sourceDoc, session);
            if (!forceRefresh && session.ParaFormatClusters != null && session.ParaFormatClusters.Count > 0)
            {
                return;
            }

            session.ParaFormatClusters = ParaFormatClusterExtractor.Extract(sourceDoc, disallowBackendApi);
        }

        public static DocumentSessionState SessionOf(Word.Document sourceDoc)
        {
            if (sourceDoc == null)
            {
                return null;
            }

            return DocumentState.GetOrCreateSession(DocumentIdentity.EnsureUuid(sourceDoc));
        }

        public static List<Dictionary<string, object>> CopyMapping(DocumentSessionState session)
        {
            if (session?.SubtypeFormatMapping == null)
            {
                return new List<Dictionary<string, object>>();
            }

            var copy = new List<Dictionary<string, object>>(session.SubtypeFormatMapping.Count);
            foreach (Dictionary<string, object> row in session.SubtypeFormatMapping)
            {
                copy.Add(row == null
                    ? null
                    : new Dictionary<string, object>(row, StringComparer.Ordinal));
            }

            return copy;
        }

        public static Dictionary<string, object> CopyDict(Dictionary<string, object> src)
        {
            if (src == null)
            {
                return null;
            }

            return new Dictionary<string, object>(src, StringComparer.Ordinal);
        }
    }
}
