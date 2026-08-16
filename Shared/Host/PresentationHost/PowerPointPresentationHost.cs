using System;
using System.IO;
using WordAddIn1.HostPlatform;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    internal static class PowerPointPresentationHost
    {
        // PpSaveAsFileType
        private const int PpSaveAsPresentation = 1;
        private const int PpSaveAsOpenXmlPresentation = 24;
        private const int PpSaveAsOpenXmlPresentationMacroEnabled = 25;

        public static bool TryOpen(
            string fullPath,
            bool createBlank,
            out OpenDocumentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (!PowerPointApplicationResolver.TryResolveTyped(
                    out PowerPoint.Application app,
                    out error,
                    createIfMissing: true))
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

            PowerPoint.Presentation presentation;
            bool created;
            bool reused = false;
            try
            {
                if (createBlank)
                {
                    presentation = app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                    presentation.SaveAs(fullPath, (PowerPoint.PpSaveAsFileType)SaveFormatFor(fullPath));
                    created = true;
                }
                else
                {
                    PowerPoint.Presentation existing = FindOpenPresentation(app, fullPath);
                    if (existing != null)
                    {
                        presentation = existing;
                        reused = true;
                    }
                    else
                    {
                        presentation = app.Presentations.Open(
                            fullPath,
                            ReadOnly: Microsoft.Office.Core.MsoTriState.msoFalse,
                            Untitled: Microsoft.Office.Core.MsoTriState.msoFalse,
                            WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);
                    }

                    created = false;
                }
            }
            catch (Exception ex)
            {
                error = "打开/新建 PowerPoint 演示文稿失败: " + ex.Message;
                return false;
            }

            PowerPointApplicationResolver.EnsureVisible(app);
            PowerPointApplicationResolver.Attach(app);
            HostCallbacks.RaisePowerPointApplicationResolved(app);
            PptChannel channel = ChannelRegistry.CreateOrGetPpt(presentation, fullPath);
            ChannelRegistry.SetDefault(channel.ChannelId);
            HostCallbacks.RaiseOpenFilesRefresh();

            result = new OpenDocumentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "ppt",
                DocUuid = channel.DocUuid,
                Path = fullPath,
                Name = channel.TryGetDisplayName() ?? Path.GetFileName(fullPath) ?? "",
                Created = created,
                Reused = reused,
                IndexReady = false
            };
            return true;
        }

        private static PowerPoint.Presentation FindOpenPresentation(
            PowerPoint.Application app,
            string fullPath)
        {
            string norm = OpenDocumentPath.NormalizePath(fullPath);
            try
            {
                foreach (PowerPoint.Presentation presentation in app.Presentations)
                {
                    try
                    {
                        string existing = PptChannel.TryReadFullName(presentation);
                        if (!string.IsNullOrEmpty(existing)
                            && string.Equals(
                                OpenDocumentPath.NormalizePath(existing),
                                norm,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            return presentation;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static int SaveFormatFor(string fullPath)
        {
            string ext = Path.GetExtension(fullPath)?.ToLowerInvariant();
            if (ext == ".ppt")
            {
                return PpSaveAsPresentation;
            }

            if (ext == ".pptm")
            {
                return PpSaveAsOpenXmlPresentationMacroEnabled;
            }

            return PpSaveAsOpenXmlPresentation;
        }
    }
}
