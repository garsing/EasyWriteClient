using System;
using System.IO;
using WordAddIn1.HostPlatform;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// Word 宿主：读内容走现网 <see cref="WordDocumentExtractor"/>。
    /// </summary>
    internal static class WordDocumentHost
    {
        public static bool TryOpen(
            string fullPath,
            bool createBlank,
            object wordApplication,
            out OpenDocumentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (!WordApplicationResolver.TryResolve(
                    wordApplication,
                    out Word.Application app,
                    out error,
                    createIfMissing: true))
            {
                return false;
            }

            try
            {
                DocumentCheckpointService.SetWordApplication(app);
                HostCallbacks.RaiseWordApplicationResolved(app);
            }
            catch (Exception)
            {
            }

            if (createBlank)
            {
                string dir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    try
                    {
                        Directory.CreateDirectory(dir);
                    }
                    catch (Exception ex)
                    {
                        error = "无法创建目录: " + ex.Message;
                        return false;
                    }
                }
            }

            Word.Document doc;
            bool created;
            bool reused = false;
            try
            {
                if (createBlank)
                {
                    doc = app.Documents.Add();
                    doc.SaveAs2(fullPath, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
                    created = true;
                }
                else
                {
                    Word.Document existing = FindOpenDocument(app, fullPath);
                    if (existing != null)
                    {
                        doc = existing;
                        reused = true;
                    }
                    else
                    {
                        doc = app.Documents.Open(fullPath);
                    }

                    created = false;
                }
            }
            catch (Exception ex)
            {
                error = "打开/新建 Word 文档失败: " + ex.Message;
                return false;
            }

            try
            {
                doc.Activate();
            }
            catch (Exception)
            {
            }

            DocumentState.BindAndActivate(doc);

            bool indexReady = false;
            string indexError = null;
            try
            {
                AgentRunCancellation.ThrowIfCancelled();
                var processing = WordDocumentExtractor.ProcessDocument(
                    doc,
                    ProcessDocumentOptions.ForGetDocumentContent("open_document"));
                indexReady = processing != null;
            }
            catch (OperationCanceledException)
            {
                error = "cancelled by user";
                return false;
            }
            catch (Exception ex)
            {
                indexError = ex.Message;
            }

            WordChannel channel = ChannelRegistry.CreateOrGetWord(doc, fullPath);
            ChannelRegistry.SetDefault(channel.ChannelId);
            HostCallbacks.RaiseOpenFilesRefresh();

            string name = null;
            try
            {
                name = doc.Name;
            }
            catch (Exception)
            {
                name = Path.GetFileName(fullPath);
            }

            result = new OpenDocumentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "word",
                DocUuid = channel.DocUuid,
                Path = fullPath,
                Name = name ?? "",
                Created = created,
                Reused = reused,
                IndexReady = indexReady,
                IndexError = indexError
            };
            return true;
        }

        private static Word.Document FindOpenDocument(Word.Application app, string fullPath)
        {
            if (app == null || string.IsNullOrEmpty(fullPath))
            {
                return null;
            }

            string norm = OpenDocumentPath.NormalizePath(fullPath);
            try
            {
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        string existing = WordChannel.TryReadFullName(d);
                        if (!string.IsNullOrEmpty(existing)
                            && string.Equals(
                                OpenDocumentPath.NormalizePath(existing),
                                norm,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            return d;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

            return null;
        }

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
