using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>将页内图片 Export 到本地目录，并回填 HTML data-src（workspace:…）。</summary>
    internal static class PptHtmlPictureExporter
    {
        // PpShapeFormatPNG = 2
        private const int PpShapeFormatPng = 2;

        public static bool TryAttachExportedPictures(
            PowerPoint.Presentation presentation,
            PptHtmlReadResult result,
            string assetsFolderName,
            string assetsLocalDir,
            out List<string> exportedRelativePaths,
            out string error)
        {
            exportedRelativePaths = new List<string>();
            error = null;
            if (presentation == null || result?.Shapes == null)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(assetsFolderName) || string.IsNullOrWhiteSpace(assetsLocalDir))
            {
                error = "assets 目录无效";
                return false;
            }

            try
            {
                Directory.CreateDirectory(assetsLocalDir);
            }
            catch (Exception ex)
            {
                error = "创建 assets 目录失败: " + ex.Message;
                return false;
            }

            if (!TryFindSlide(presentation, result.SlideId, out PowerPoint.Slide slide, out error))
            {
                return false;
            }

            foreach (PptHtmlShapeNode node in result.Shapes)
            {
                if (node == null || !string.Equals(node.ShapeType, "picture", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!PptShapeId.TryParseShape(node.ShapeId, out _, out int comId))
                {
                    continue;
                }

                PowerPoint.Shape shape = FindShapeById(slide.Shapes, comId);
                if (shape == null)
                {
                    continue;
                }

                string fileName = node.ShapeId + ".png";
                string safeFile = WorkspacePathResolver.SanitizeFilename(fileName);
                if (safeFile == null)
                {
                    safeFile = "s" + comId.ToString(CultureInfo.InvariantCulture) + ".png";
                }

                string localPath = Path.Combine(assetsLocalDir, safeFile);
                try
                {
                    if (File.Exists(localPath))
                    {
                        File.Delete(localPath);
                    }

                    shape.Export(localPath, (PowerPoint.PpShapeFormat)PpShapeFormatPng);
                }
                catch (Exception ex)
                {
                    error = "导出图片失败 (" + node.ShapeId + "): " + ex.Message;
                    return false;
                }

                if (!File.Exists(localPath))
                {
                    error = "导出图片未生成文件: " + node.ShapeId;
                    return false;
                }

                string relative = assetsFolderName.Trim().TrimEnd('/', '\\') + "/" + safeFile;
                relative = WorkspacePathResolver.SanitizeWorkspaceRelativePath(relative);
                if (relative == null)
                {
                    error = "非法 assets 相对路径: " + assetsFolderName + "/" + safeFile;
                    return false;
                }

                node.DataSrc = "workspace:" + relative;
                node.Editable = true;
                exportedRelativePaths.Add(relative);
            }

            return true;
        }

        private static bool TryFindSlide(
            PowerPoint.Presentation presentation,
            string slideIdText,
            out PowerPoint.Slide slide,
            out string error)
        {
            slide = null;
            error = null;
            if (!int.TryParse(slideIdText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int slideId))
            {
                error = "slide_id 无效";
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

        private static PowerPoint.Shape FindShapeById(PowerPoint.Shapes shapes, int id)
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
                            PowerPoint.Shape found = FindInGroup(shape, id);
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

        private static PowerPoint.Shape FindInGroup(PowerPoint.Shape group, int id)
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
                        PowerPoint.Shape nested = FindInGroup(child, id);
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
    }
}
