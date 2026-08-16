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
                // 先查找已打开稿，再取 Presentations 集合做 Open/Add。
                // 禁止在查找时 FinalRelease 演示文稿 RCW（会拆掉后续 Open 所用的 COM）。
                if (createBlank)
                {
                    object presentations = WppCom.GetProperty(app, "Presentations");
                    presentation = TryAddPresentation(presentations);
                    if (presentation == null)
                    {
                        error = "WPS 演示新建失败（Presentations.Add 无返回）";
                        return false;
                    }

                    TrySaveAs(presentation, fullPath);
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
                        object presentations = WppCom.GetProperty(app, "Presentations");
                        presentation = TryOpenPresentation(presentations, fullPath);
                        if (presentation == null)
                        {
                            error = "WPS 演示打开失败（Presentations.Open 无返回）";
                            return false;
                        }
                    }

                    created = false;
                }
            }
            catch (Exception ex)
            {
                error = "打开/新建 WPS 演示失败: " + ex.Message;
                return false;
            }

            string displayName = WppCom.TryReadName(presentation) ?? Path.GetFileName(fullPath) ?? "";
            WppCom.EnsureVisible(app, displayName);
            WppChannel channel = ChannelRegistry.CreateOrGetWpp(presentation, fullPath);
            ChannelRegistry.SetDefault(channel.ChannelId);

            result = new OpenDocumentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "wpp",
                DocUuid = channel.DocUuid,
                Path = fullPath,
                Name = displayName,
                Created = created,
                Reused = reused,
                IndexReady = false
            };

            // 建渠完成后再刷新侧栏，避免探测 Snapshot 抢先 FinalRelease 打开路径持有的 RCW
            HostCallbacks.RaiseOpenFilesRefresh();
            return true;
        }

        /// <summary>对齐 Et：查找时不 ComRelease 枚举到的对象。</summary>
        private static object FindOpenPresentation(object app, string fullPath)
        {
            string norm = OpenDocumentPath.NormalizePath(fullPath);
            foreach (object presentation in WppCom.EnumeratePresentations(app))
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

            return null;
        }

        private static object TryAddPresentation(object presentations)
        {
            if (presentations == null)
            {
                return null;
            }

            try
            {
                // PowerPoint 风格：Add(WithWindow)
                return WppCom.Invoke(presentations, "Add", true);
            }
            catch (Exception)
            {
            }

            try
            {
                return WppCom.Invoke(presentations, "Add");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryOpenPresentation(object presentations, string fullPath)
        {
            if (presentations == null)
            {
                return null;
            }

            // 1) 仅路径（对齐 WPS 文字 Documents.Open）
            try
            {
                object opened = WppCom.Invoke(presentations, "Open", fullPath);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception)
            {
            }

            // 2) Open(FileName, ReadOnly, Untitled, WithWindow) — 对齐 PowerPoint
            try
            {
                object opened = WppCom.Invoke(presentations, "Open", fullPath, false, false, true);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception)
            {
            }

            // 3) 部分 WPS：Open(FileName, WithWindow)
            try
            {
                return WppCom.Invoke(presentations, "Open", fullPath, true);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TrySaveAs(object presentation, string fullPath)
        {
            try
            {
                WppCom.Invoke(presentation, "SaveAs", fullPath);
            }
            catch (Exception)
            {
                // 部分版本 SaveAs 第二参数为格式枚举；无格式再试一次仅路径
                WppCom.Invoke(presentation, "SaveAs", fullPath, 1);
            }
        }
    }
}
