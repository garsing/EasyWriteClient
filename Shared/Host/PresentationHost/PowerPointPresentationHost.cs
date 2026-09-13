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
            // 临时关闭：概览扫全稿颜色太慢，需要时再打开。
            // try
            // {
            //     PptPaletteCollectResult palette = PptPaletteIo.CollectPowerPoint(presentation);
            //     result.Palette = palette.Palette;
            //     result.PaletteSampled = palette.Sampled;
            //     result.PaletteScannedCount = palette.ScannedCount;
            // }
            // catch (Exception)
            // {
            //     result.Palette = new List<string> { "#000000", "#FFFFFF", "none" };
            // }

            return true;
        }

        public static bool TryCaptureSlide(
            PptChannel channel,
            int pageNumber,
            out PresentationCaptureResult result,
            out string error,
            int maxLongEdge = 0)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            int count;
            try
            {
                count = presentation.Slides.Count;
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            if (pageNumber < 1 || pageNumber > count)
            {
                error = "页码 " + pageNumber + " 超出范围，演示文稿共 " + count + " 页";
                return false;
            }

            PowerPoint.Slide slide;
            try
            {
                slide = presentation.Slides[pageNumber];
                if (slide.SlideIndex != pageNumber)
                {
                    slide = null;
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Slide candidate = presentation.Slides[i];
                        if (candidate.SlideIndex == pageNumber)
                        {
                            slide = candidate;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = "定位幻灯片失败: " + ex.Message;
                return false;
            }

            if (slide == null)
            {
                error = "页码 " + pageNumber + " 超出范围，演示文稿共 " + count + " 页";
                return false;
            }

            string tempDir = Path.Combine(Path.GetTempPath(), "EasyWrite", "capture", Guid.NewGuid().ToString("N"));
            string pngPath = Path.Combine(tempDir, "slide.png");
            try
            {
                Directory.CreateDirectory(tempDir);
                float slideWidth = 0f;
                float slideHeight = 0f;
                try
                {
                    slideWidth = presentation.PageSetup.SlideWidth;
                    slideHeight = presentation.PageSetup.SlideHeight;
                }
                catch (Exception)
                {
                }

                ImageCaptureCompressor.FitExportPixelSize(
                    slideWidth,
                    slideHeight,
                    maxLongEdge,
                    out int exportW,
                    out int exportH);
                slide.Export(pngPath, "PNG", exportW, exportH);
                if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                {
                    error = "幻灯片导出图片为空";
                    return false;
                }

                string slideId = "";
                try
                {
                    slideId = Convert.ToString(slide.SlideID) ?? "";
                }
                catch (Exception)
                {
                }

                result = new PresentationCaptureResult
                {
                    ChannelId = channel.ChannelId,
                    Kind = "ppt",
                    PageNumber = pageNumber,
                    SlideCount = count,
                    SlideId = slideId,
                    Image = ImageCaptureCompressor.CompressFile(pngPath)
                };
                return true;
            }
            catch (Exception ex)
            {
                error = "导出幻灯片失败: " + ex.Message;
                return false;
            }
            finally
            {
                try
                {
                    if (File.Exists(pngPath))
                    {
                        File.Delete(pngPath);
                    }

                    if (Directory.Exists(tempDir))
                    {
                        Directory.Delete(tempDir, true);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        public static bool TryCaptureShape(
            PptChannel channel,
            string shapeId,
            out PresentationCaptureResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (!PptShapeId.TryParseShape(shapeId, out string slideIdText, out int comId))
            {
                error = "形状截图须提供完整 shape_id（sid{SlideID}-s{n}）";
                return false;
            }

            if (!TryFindSlideBySlideId(presentation, slideIdText, out PowerPoint.Slide slide, out error))
            {
                return false;
            }

            PowerPoint.Shape shape = FindShapeByIdDeep(slide.Shapes, comId);
            if (shape == null)
            {
                error = "未找到形状: " + shapeId.Trim();
                return false;
            }

            try
            {
                if (shape.Visible == Office.MsoTriState.msoFalse)
                {
                    error = "形状已隐藏，无法截图";
                    return false;
                }
            }
            catch (Exception)
            {
            }

            float shapeLeft;
            float shapeTop;
            float shapeWidth;
            float shapeHeight;
            try
            {
                shapeLeft = shape.Left;
                shapeTop = shape.Top;
                shapeWidth = shape.Width;
                shapeHeight = shape.Height;
            }
            catch (Exception ex)
            {
                error = "读取形状几何失败: " + ex.Message;
                return false;
            }

            int slideCount;
            int pageNumber;
            try
            {
                slideCount = presentation.Slides.Count;
                pageNumber = slide.SlideIndex;
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            string tempDir = Path.Combine(Path.GetTempPath(), "EasyWrite", "capture", Guid.NewGuid().ToString("N"));
            string pngPath = Path.Combine(tempDir, "slide.png");
            try
            {
                Directory.CreateDirectory(tempDir);
                float slideWidth = 0f;
                float slideHeight = 0f;
                try
                {
                    slideWidth = presentation.PageSetup.SlideWidth;
                    slideHeight = presentation.PageSetup.SlideHeight;
                }
                catch (Exception)
                {
                }

                ImageCaptureCompressor.FitExportPixelSize(
                    slideWidth,
                    slideHeight,
                    out int exportW,
                    out int exportH);
                slide.Export(pngPath, "PNG", exportW, exportH);
                if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                {
                    error = "幻灯片导出图片为空";
                    return false;
                }

                if (!PptShapeCaptureHelper.TryCropFromSlidePng(
                        pngPath,
                        slideWidth,
                        slideHeight,
                        shapeLeft,
                        shapeTop,
                        shapeWidth,
                        shapeHeight,
                        out PptShapeCaptureHelper.CropResult crop,
                        out error))
                {
                    return false;
                }

                try
                {
                    result = new PresentationCaptureResult
                    {
                        ChannelId = channel.ChannelId,
                        Kind = "ppt",
                        PageNumber = pageNumber,
                        SlideCount = slideCount,
                        SlideId = slideIdText,
                        ShapeId = shapeId.Trim(),
                        Scale = crop.Scale,
                        BoundsLeftPct = crop.LeftPct,
                        BoundsTopPct = crop.TopPct,
                        BoundsWidthPct = crop.WidthPct,
                        BoundsHeightPct = crop.HeightPct,
                        Image = ImageCaptureCompressor.Compress(crop.Bitmap)
                    };
                    return true;
                }
                finally
                {
                    if (crop?.Bitmap != null)
                    {
                        crop.Bitmap.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                error = "导出形状截图失败: " + ex.Message;
                return false;
            }
            finally
            {
                try
                {
                    if (File.Exists(pngPath))
                    {
                        File.Delete(pngPath);
                    }

                    if (Directory.Exists(tempDir))
                    {
                        Directory.Delete(tempDir, true);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool TryFindSlideBySlideId(
            PowerPoint.Presentation presentation,
            string slideIdText,
            out PowerPoint.Slide slide,
            out string error)
        {
            slide = null;
            error = null;
            if (!int.TryParse(slideIdText, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out int slideId))
            {
                error = "shape_id 内 SlideID 无效";
                return false;
            }

            try
            {
                slide = presentation.Slides.FindBySlideID(slideId);
                if (slide != null)
                {
                    return true;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                int count = presentation.Slides.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Slide s = presentation.Slides[i];
                    if (s.SlideID == slideId)
                    {
                        slide = s;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                error = "查找幻灯片失败: " + ex.Message;
                return false;
            }

            error = "幻灯片不存在: " + slideIdText;
            return false;
        }

        private static PowerPoint.Shape FindShapeByIdDeep(PowerPoint.Shapes shapes, int id)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape = shapes[i];
                    try
                    {
                        if (shape.Id == id)
                        {
                            return shape;
                        }

                        if ((int)shape.Type == 6)
                        {
                            PowerPoint.Shape found = FindShapeInGroupDeep(shape, id);
                            if (found != null)
                            {
                                return found;
                            }
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

        private static PowerPoint.Shape FindShapeInGroupDeep(PowerPoint.Shape group, int id)
        {
            try
            {
                PowerPoint.GroupShapes items = group.GroupItems;
                int count = items.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape child = items[i];
                    if (child.Id == id)
                    {
                        return child;
                    }

                    if ((int)child.Type == 6)
                    {
                        PowerPoint.Shape nested = FindShapeInGroupDeep(child, id);
                        if (nested != null)
                        {
                            return nested;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        public static bool TryReadPptHtml(
            PptChannel channel,
            string slideId,
            out PptHtmlReadResult result,
            out string error)
        {
            return TryReadPptHtml(channel, slideId, null, out result, out error);
        }

        public static bool TryReadPptHtml(
            PptChannel channel,
            string slideId,
            string shapeId,
            out PptHtmlReadResult result,
            out string error)
        {
            return TryReadPptHtml(channel, slideId, shapeId, false, out result, out error);
        }

        public static bool TryReadPptHtml(
            PptChannel channel,
            string slideId,
            string shapeId,
            bool fullPage,
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
                shapeId,
                fullPage,
                out result,
                out error);
        }

        public static bool TrySearchPptHtml(
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
                null,
                false,
                true,
                out result,
                out error);
        }

        public static bool TryExportPptHtmlPictures(
            PptChannel channel,
            PptHtmlReadResult result,
            string assetsFolderName,
            string assetsLocalDir,
            out System.Collections.Generic.List<string> exportedRelativePaths,
            out string error)
        {
            exportedRelativePaths = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptHtmlPictureExporter.TryAttachExportedPictures(
                presentation,
                result,
                assetsFolderName,
                assetsLocalDir,
                out exportedRelativePaths,
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
            PptChannel destChannel,
            PptChannel sourceChannel,
            PresentationManageSlideRequest request,
            out PresentationManageSlideResult result,
            out string error)
        {
            result = null;
            error = null;
            if (destChannel == null || !destChannel.TryGetLivePresentation(out PowerPoint.Presentation dest))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (sourceChannel == null || !sourceChannel.TryGetLivePresentation(out PowerPoint.Presentation source))
            {
                error = "源头演示文稿已关闭";
                return false;
            }

            return PptSlidePowerPointManager.TryManage(
                dest,
                source,
                destChannel.ChannelId,
                sourceChannel.ChannelId,
                request,
                out result,
                out error);
        }

        public static bool TryManageShape(
            PptChannel destChannel,
            PptChannel sourceChannel,
            PresentationManageShapeRequest request,
            out PresentationManageShapeResult result,
            out string error)
        {
            result = null;
            error = null;
            if (destChannel == null || !destChannel.TryGetLivePresentation(out PowerPoint.Presentation dest))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (sourceChannel == null || !sourceChannel.TryGetLivePresentation(out PowerPoint.Presentation source))
            {
                error = "源头演示文稿已关闭";
                return false;
            }

            return PptShapePowerPointManager.TryManage(
                dest,
                source,
                destChannel.ChannelId,
                sourceChannel.ChannelId,
                request,
                out result,
                out error);
        }

        public static bool TryPptAnimation(
            PptChannel channel,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptAnimationPowerPointManager.TryAnimate(
                presentation,
                channel.ChannelId,
                request,
                out result,
                out error);
        }

        public static bool TryPptTransition(
            PptChannel channel,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            return PptTransitionPowerPointManager.TryTransition(
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
                SlideId = ""
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

            return info;
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
