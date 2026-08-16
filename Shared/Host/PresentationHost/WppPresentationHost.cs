using System;
using System.IO;
using WordAddIn1.HostPlatform;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class WppPresentationHost
    {
        public static bool TryOpen(
            string fullPath,
            bool createBlank,
            out OpenDocumentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (!WppApplicationResolver.TryResolve(out object app, out error, createIfMissing: true))
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

            object presentation;
            bool created;
            bool reused = false;
            try
            {
                object presentations = WppCom.GetProperty(app, "Presentations");
                if (createBlank)
                {
                    presentation = WppCom.Invoke(presentations, "Add");
                    WppCom.Invoke(presentation, "SaveAs", fullPath);
                    created = true;
                }
                else
                {
                    object existing = FindOpenPresentation(app, fullPath);
                    if (existing != null)
                    {
                        presentation = existing;
                        reused = true;
                    }
                    else
                    {
                        presentation = WppCom.Invoke(presentations, "Open", fullPath);
                    }

                    created = false;
                }
            }
            catch (Exception ex)
            {
                error = "打开/新建 WPS 演示失败: " + ex.Message;
                return false;
            }

            WppCom.EnsureVisible(app, Path.GetFileName(fullPath));
            WppChannel channel = ChannelRegistry.CreateOrGetWpp(presentation, fullPath);
            ChannelRegistry.SetDefault(channel.ChannelId);
            HostCallbacks.RaiseOpenFilesRefresh();

            result = new OpenDocumentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "wpp",
                DocUuid = channel.DocUuid,
                Path = fullPath,
                Name = WppCom.TryReadName(presentation) ?? Path.GetFileName(fullPath) ?? "",
                Created = created,
                Reused = reused,
                IndexReady = false
            };
            return true;
        }

        private static object FindOpenPresentation(object app, string fullPath)
        {
            string norm = OpenDocumentPath.NormalizePath(fullPath);
            foreach (object presentation in WppCom.EnumeratePresentations(app))
            {
                try
                {
                    string existing = WppCom.TryReadFullName(presentation);
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

                ComRelease.Safe(presentation);
            }

            return null;
        }
    }
}
