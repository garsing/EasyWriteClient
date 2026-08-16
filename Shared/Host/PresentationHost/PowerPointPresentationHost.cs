using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1.HostPlatform;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PowerPointPresentationHost
    {
        private const int PpSaveAsPresentation = 1;
        private const int PpSaveAsOpenXmlPresentation = 24;
        private const int PpSaveAsOpenXmlPresentationMacroEnabled = 25;

        private const int PpPlaceholderTitle = 1;
        private const int PpPlaceholderCenterTitle = 3;
        private const int PpPlaceholderVerticalTitle = 16;
        private const int PpPlaceholderBody = 2;
        private const int PpPlaceholderVerticalBody = 17;

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
                    presentation = app.Presentations.Add(Office.MsoTriState.msoTrue);
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
                            ReadOnly: Office.MsoTriState.msoFalse,
                            Untitled: Office.MsoTriState.msoFalse,
                            WithWindow: Office.MsoTriState.msoTrue);
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

        public static bool TryGetPresentationContent(
            PptChannel channel,
            out PresentationContentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            string name;
            string path;
            int count;
            try
            {
                name = presentation.Name ?? "";
                path = NormalizeSavedPath(PptChannel.TryReadFullName(presentation));
                count = presentation.Slides.Count;
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
                    PowerPoint.Slide slide = presentation.Slides[i];
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
                Kind = "ppt",
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
            PptChannel channel,
            string slideId,
            out PptHtmlReadResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptHtmlPowerPointReader.TryRead(
                presentation,
                slideId,
                channel.ChannelId,
                "ppt",
                out result,
                out error);
        }

        public static bool TryApplyPptHtml(
            PptChannel channel,
            PptHtmlApplyPlan plan,
            out PptHtmlApplyResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptHtmlPowerPointApplier.TryApply(
                presentation,
                plan,
                channel.ChannelId,
                out result,
                out error);
        }

        public static bool TryManageSlide(
            PptChannel channel,
            PresentationManageSlideRequest request,
            out PresentationManageSlideResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptSlidePowerPointManager.TryManage(
                presentation,
                channel.ChannelId,
                request,
                out result,
                out error);
        }

        private static PresentationSlideInfo ReadSlide(PowerPoint.Slide slide, int fallbackIndex)
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
                info.Index = slide.SlideIndex;
            }
            catch (Exception)
            {
            }

            try
            {
                info.SlideId = Convert.ToString(slide.SlideID) ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                info.Hidden = slide.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
            }

            try
            {
                info.Layout = slide.CustomLayout != null ? (slide.CustomLayout.Name ?? "") : "";
            }
            catch (Exception)
            {
                try
                {
                    info.Layout = Convert.ToString(slide.Layout) ?? "";
                }
                catch (Exception)
                {
                }
            }

            info.Title = TryReadTitle(slide) ?? "";
            info.HasNotes = TryDetectHasNotes(slide);
            return info;
        }

        private static string TryReadTitle(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return "";
            }

            try
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        {
                            continue;
                        }

                        int ph = Convert.ToInt32(shape.PlaceholderFormat.Type);
                        if (ph != PpPlaceholderTitle
                            && ph != PpPlaceholderCenterTitle
                            && ph != PpPlaceholderVerticalTitle)
                        {
                            continue;
                        }

                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                        {
                            continue;
                        }

                        string text = shape.TextFrame.TextRange.Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            return text.Replace("\r", "\n").Trim();
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

        private static bool? TryDetectHasNotes(PowerPoint.Slide slide)
        {
            try
            {
                PowerPoint.SlideRange notes = slide.NotesPage;
                if (notes == null)
                {
                    return false;
                }

                foreach (PowerPoint.Shape shape in notes.Shapes)
                {
                    try
                    {
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            int ph = Convert.ToInt32(shape.PlaceholderFormat.Type);
                            if (ph != PpPlaceholderBody && ph != PpPlaceholderVerticalBody)
                            {
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }

                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                        {
                            continue;
                        }

                        string text = shape.TextFrame.TextRange.Text;
                        if (!string.IsNullOrWhiteSpace(text))
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
