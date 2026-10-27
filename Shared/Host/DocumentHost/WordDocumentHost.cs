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
            using (NativeWindowActivate.MaybeQuietOfficeLaunch())
            {
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
                    if (!TrySaveNewDocument(app, doc, fullPath, out error))
                    {
                        TryCloseWithoutSave(doc);
                        return false;
                    }

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

            DocumentState.BindAndActivate(doc);

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
                IndexReady = false
            };
            return true;
            }
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

        /// <summary>
        /// 建图后收无窗 Excel 时，下一份空白稿的 SaveAs 会偶发「文件许可权」。
        /// 退出图表编辑后重试，失败则关掉这次没存上的空稿，避免留着 DocumentN。
        /// </summary>
        private static bool TrySaveNewDocument(
            Word.Application app,
            Word.Document doc,
            string fullPath,
            out string error)
        {
            error = null;
            const int attempts = 3;
            for (int i = 0; i < attempts; i++)
            {
                try
                {
                    ChartComUiHelper.TryExitChartDataEditMode(app);
                    doc.SaveAs2(fullPath, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
                    return true;
                }
                catch (Exception ex)
                {
                    error = "打开/新建 Word 文档失败: " + ex.Message;
                    bool permission = ex.Message != null
                        && (ex.Message.IndexOf("许可权", StringComparison.Ordinal) >= 0
                            || ex.Message.IndexOf("permission", StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!permission || i == attempts - 1)
                    {
                        return false;
                    }

                    System.Threading.Thread.Sleep(500);
                }
            }

            return false;
        }

        private static void TryCloseWithoutSave(Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            try
            {
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch (Exception)
            {
            }
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
