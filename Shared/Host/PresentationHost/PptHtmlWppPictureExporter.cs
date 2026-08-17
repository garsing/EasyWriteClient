using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>WPP：晚绑定 Export 页内图片到本地，并回填 data-src。</summary>
    internal static class PptHtmlWppPictureExporter
    {
        private const int PpShapeFormatPng = 2;

        public static bool TryAttachExportedPictures(
            object presentation,
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

            if (!TryFindSlide(presentation, result.SlideId, out object slide, out error))
            {
                return false;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            if (shapes == null)
            {
                return true;
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

                object shape = FindShapeById(shapes, comId);
                if (shape == null)
                {
                    continue;
                }

                string tempPath = Path.Combine(
                    assetsLocalDir,
                    "_tmp_" + Guid.NewGuid().ToString("N") + ".png");
                try
                {
                    WppCom.Invoke(shape, "Export", tempPath, PpShapeFormatPng);
                }
                catch (Exception ex)
                {
                    TryDeleteQuiet(tempPath);
                    error = "导出图片失败 (" + node.ShapeId + "): " + ex.Message;
                    return false;
                }

                if (!File.Exists(tempPath))
                {
                    error = "导出图片未生成文件: " + node.ShapeId;
                    return false;
                }

                if (!PptHtmlImageFileNaming.TryFinalizeExportedPng(
                        tempPath,
                        assetsFolderName,
                        assetsLocalDir,
                        out string relative,
                        out error))
                {
                    return false;
                }

                node.DataSrc = "workspace:" + relative;
                node.Editable = true;
                if (!exportedRelativePaths.Contains(relative))
                {
                    exportedRelativePaths.Add(relative);
                }
            }

            return true;
        }

        private static void TryDeleteQuiet(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryFindSlide(object presentation, string slideIdText, out object slide, out string error)
        {
            slide = null;
            error = null;
            if (!int.TryParse(slideIdText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int slideId))
            {
                error = "slide_id 无效";
                return false;
            }

            object slides = WppCom.GetProperty(presentation, "Slides");
            if (slides == null)
            {
                error = "无法访问 Slides";
                return false;
            }

            try
            {
                slide = WppCom.Invoke(slides, "FindBySlideID", slideId);
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
                int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object s = WppCom.GetIndexed(slides, i);
                    if (s == null)
                    {
                        continue;
                    }

                    object idObj = WppCom.GetProperty(s, "SlideID");
                    if (idObj != null && Convert.ToInt32(idObj) == slideId)
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

        private static object FindShapeById(object shapes, int id)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object shape = WppCom.GetIndexed(shapes, i);
                    if (shape == null)
                    {
                        continue;
                    }

                    object idObj = WppCom.GetProperty(shape, "Id");
                    if (idObj != null && Convert.ToInt32(idObj) == id)
                    {
                        return shape;
                    }

                    object typeObj = WppCom.GetProperty(shape, "Type");
                    if (typeObj != null && Convert.ToInt32(typeObj) == 6)
                    {
                        object found = FindInGroup(shape, id);
                        if (found != null)
                        {
                            return found;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static object FindInGroup(object group, int id)
        {
            try
            {
                object items = WppCom.GetProperty(group, "GroupItems");
                if (items == null)
                {
                    return null;
                }

                int count = Convert.ToInt32(WppCom.GetProperty(items, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object child = WppCom.GetIndexed(items, i);
                    if (child == null)
                    {
                        continue;
                    }

                    object idObj = WppCom.GetProperty(child, "Id");
                    if (idObj != null && Convert.ToInt32(idObj) == id)
                    {
                        return child;
                    }

                    object typeObj = WppCom.GetProperty(child, "Type");
                    if (typeObj != null && Convert.ToInt32(typeObj) == 6)
                    {
                        object nested = FindInGroup(child, id);
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
