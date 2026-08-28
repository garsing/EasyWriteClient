using System;
using System.IO;
using WordAddIn1.HostPlatform;
using WordAddIn1.OpenFiles;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// WPS 宿主：晚绑定文档；优先兼容复用 Word 抽取管线（渠道仍为 wps:）。
    /// 后端 API 分块本期禁止（I6b）。
    /// </summary>
    internal static class WpsDocumentHost
    {
        public static bool TryOpen(
            string fullPath,
            bool createBlank,
            out OpenDocumentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (!WpsApplicationResolver.TryResolve(out object app, out error, createIfMissing: true))
            {
                return false;
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

            object doc;
            bool created;
            bool reused = false;
            try
            {
                object documents = WpsCom.GetProperty(app, "Documents");
                if (createBlank)
                {
                    doc = WpsCom.Invoke(documents, "Add");
                    WpsCom.Invoke(doc, "SaveAs", fullPath);
                    created = true;
                }
                else
                {
                    object existing = FindOpenDocument(app, fullPath);
                    if (existing != null)
                    {
                        doc = existing;
                        reused = true;
                    }
                    else
                    {
                        doc = WpsCom.Invoke(documents, "Open", fullPath);
                    }

                    created = false;
                }
            }
            catch (Exception ex)
            {
                error = "打开/新建 WPS 文字文档失败: " + ex.Message;
                return false;
            }

            bool indexReady = false;
            string indexError = null;
            try
            {
                var context = new DocumentSessionContext("", ChannelKind.Wps, "", null, doc);
                if (DocumentHostCompat.TryPrepareWpsInteropDocument(
                        context,
                        out Word.Document wordDoc,
                        out _,
                        activateDocument: false)
                    && wordDoc != null)
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    var processing = WordDocumentExtractor.ProcessDocument(
                        wordDoc,
                        ProcessDocumentOptions.ForGetDocumentContent("open_document"));
                    indexReady = processing != null;
                }
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

            WpsChannel channel = ChannelRegistry.CreateOrGetWps(doc, fullPath);
            ChannelRegistry.SetDefault(channel.ChannelId);
            HostCallbacks.RaiseOpenFilesRefresh();

            result = new OpenDocumentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "wps",
                DocUuid = channel.DocUuid,
                Path = fullPath,
                Name = WpsCom.TryReadName(doc) ?? Path.GetFileName(fullPath) ?? "",
                Created = created,
                Reused = reused,
                IndexReady = indexReady,
                IndexError = indexError
            };
            return true;
        }

        private static object FindOpenDocument(object app, string fullPath)
        {
            string norm = OpenDocumentPath.NormalizePath(fullPath);
            foreach (object doc in WpsCom.EnumerateDocuments(app))
            {
                string existing = WpsCom.TryReadFullName(doc);
                if (!string.IsNullOrEmpty(existing)
                    && string.Equals(
                        OpenDocumentPath.NormalizePath(existing),
                        norm,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }
            }

            return null;
        }

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
