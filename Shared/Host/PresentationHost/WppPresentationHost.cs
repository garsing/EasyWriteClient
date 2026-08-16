using System;
using System.Collections.Generic;
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

        public static bool TryGetPresentationContent(
            WppChannel channel,
            out PresentationContentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out object presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            string name;
            string path;
            int count;
            object slidesCollection;
            try
            {
                name = WppCom.TryReadName(presentation) ?? "";
                path = NormalizeSavedPath(WppCom.TryReadFullName(presentation));
                slidesCollection = WppCom.GetProperty(presentation, "Slides");
                count = Convert.ToInt32(WppCom.GetProperty(slidesCollection, "Count"));
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            var slides = new List<PresentationSlideInfo>();
            bool truncated = false;
            string truncatedReason = null;
            int limit = Math.Min(count, PresentationContentResult.MaxSlides);
            if (count > PresentationContentResult.MaxSlides)
            {
                truncated = true;
                truncatedReason = "页数超过 " + PresentationContentResult.MaxSlides
                    + "，仅返回前 " + PresentationContentResult.MaxSlides + " 页";
            }

            try
            {
                for (int i = 1; i <= limit; i++)
                {
                    object slide = WppCom.GetIndexed(slidesCollection, i);
                    if (slide == null)
                    {
                        continue;
                    }

                    slides.Add(ReadSlide(slide, i));
                }
            }
            catch (Exception ex)
            {
                error = "读取幻灯片列表失败: " + ex.Message;
                return false;
            }

            result = new PresentationContentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "wpp",
                Name = name,
                Path = path ?? "",
                SlideCount = count,
                Slides = slides,
                Truncated = truncated,
                TruncatedReason = truncatedReason
            };
            return true;
        }

        public static bool TryReadPptHtml(
            WppChannel channel,
            string slideId,
            out PptHtmlReadResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out object presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptHtmlWppReader.TryRead(
                presentation,
                slideId,
                channel.ChannelId,
                "wpp",
                out result,
                out error);
        }

        public static bool TryApplyPptHtml(
            WppChannel channel,
            PptHtmlApplyPlan plan,
            out PptHtmlApplyResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out object presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptHtmlWppApplier.TryApply(
                presentation,
                plan,
                channel.ChannelId,
                out result,
                out error);
        }

        private static PresentationSlideInfo ReadSlide(object slide, int fallbackIndex)
        {
            var info = new PresentationSlideInfo
            {
                Index = fallbackIndex,
                SlideId = "",
                Title = "",
                Layout = "",
                Hidden = false,
                HasNotes = null
            };

            try
            {
                object index = WppCom.GetProperty(slide, "SlideIndex");
                if (index != null)
                {
                    info.Index = Convert.ToInt32(index);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object id = WppCom.GetProperty(slide, "SlideID");
                info.SlideId = id == null ? "" : Convert.ToString(id) ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                object transition = WppCom.GetProperty(slide, "SlideShowTransition");
                object hidden = WppCom.GetProperty(transition, "Hidden");
                info.Hidden = IsTruthyMso(hidden);
            }
            catch (Exception)
            {
            }

            try
            {
                object layout = WppCom.GetProperty(slide, "CustomLayout");
                object layoutName = WppCom.GetProperty(layout, "Name");
                info.Layout = layoutName == null ? "" : Convert.ToString(layoutName) ?? "";
            }
            catch (Exception)
            {
            }

            info.Title = TryReadTitle(slide) ?? "";
            info.HasNotes = TryDetectHasNotes(slide);
            return info;
        }

        private static string TryReadTitle(object slide)
        {
            try
            {
                object shapes = WppCom.GetProperty(slide, "Shapes");
                int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object shape = WppCom.GetIndexed(shapes, i);
                    if (shape == null)
                    {
                        continue;
                    }

                    try
                    {
                        object ph = WppCom.GetProperty(shape, "PlaceholderFormat");
                        if (ph == null)
                        {
                            continue;
                        }

                        int phType = Convert.ToInt32(WppCom.GetProperty(ph, "Type"));
                        // title / center title / vertical title
                        if (phType != 1 && phType != 3 && phType != 16)
                        {
                            continue;
                        }

                        object hasTf = WppCom.GetProperty(shape, "HasTextFrame");
                        if (!IsTruthyMso(hasTf))
                        {
                            continue;
                        }

                        object tf = WppCom.GetProperty(shape, "TextFrame");
                        object tr = WppCom.GetProperty(tf, "TextRange");
                        object text = WppCom.GetProperty(tr, "Text");
                        string s = text == null ? "" : Convert.ToString(text) ?? "";
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            return s.Replace("\r", "\n").Trim();
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

            return "";
        }

        private static bool? TryDetectHasNotes(object slide)
        {
            try
            {
                object notes = WppCom.GetProperty(slide, "NotesPage");
                if (notes == null)
                {
                    return false;
                }

                object shapes = WppCom.GetProperty(notes, "Shapes");
                int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object shape = WppCom.GetIndexed(shapes, i);
                    if (shape == null)
                    {
                        continue;
                    }

                    try
                    {
                        object ph = WppCom.GetProperty(shape, "PlaceholderFormat");
                        if (ph == null)
                        {
                            continue;
                        }

                        int phType = Convert.ToInt32(WppCom.GetProperty(ph, "Type"));
                        // body / vertical body
                        if (phType != 2 && phType != 17)
                        {
                            continue;
                        }

                        object hasTf = WppCom.GetProperty(shape, "HasTextFrame");
                        if (!IsTruthyMso(hasTf))
                        {
                            continue;
                        }

                        object tf = WppCom.GetProperty(shape, "TextFrame");
                        object tr = WppCom.GetProperty(tf, "TextRange");
                        object text = WppCom.GetProperty(tr, "Text");
                        string s = text == null ? "" : Convert.ToString(text) ?? "";
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool IsTruthyMso(object value)
        {
            if (value == null)
            {
                return false;
            }

            try
            {
                if (value is bool b)
                {
                    return b;
                }

                int n = Convert.ToInt32(value);
                return n == -1 || n == 1;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string NormalizeSavedPath(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return "";
            }

            try
            {
                if (!Path.IsPathRooted(fullName))
                {
                    return "";
                }

                return Path.GetFullPath(fullName);
            }
            catch (Exception)
            {
                return fullName;
            }
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
