using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class PptLibExtract
    {
        public static bool TryExtract(
            IOperationChannel channel,
            string slideId,
            string shapeId,
            string name,
            string description,
            string category,
            string userId,
            bool overwrite,
            out Dictionary<string, object> data,
            out ToolResult errorResult)
        {
            data = null;
            errorResult = null;
            if (!PptShapeId.TryParseShapeComId(shapeId, out int comId))
            {
                errorResult = new ToolResult { Success = false, Error = "非法 ShapeId: " + (shapeId ?? "") };
                return false;
            }

            if (!TryCollectIds(
                    channel,
                    slideId,
                    comId,
                    out _,
                    out object foundShape,
                    out int pageIndex,
                    out string collectError))
            {
                errorResult = new ToolResult { Success = false, Error = collectError };
                return false;
            }

            if (!TryReadExtractTree(channel, slideId, comId, out PptHtmlShapeNode tree, out string readError))
            {
                errorResult = new ToolResult { Success = false, Error = readError };
                return false;
            }

            StripShapeIds(tree);
            string fragment = PptConventionHtml.BuildFragment(new List<PptHtmlShapeNode> { tree });
            TryReadComSlideBox(channel, slideId, foundShape, out double boxL, out double boxT, out double boxW, out double boxH);
            var slots = new List<string>();
            var meta = new JObject
            {
                ["id"] = userId,
                ["name"] = name,
                ["description"] = description,
                ["category"] = category,
                ["source"] = "user",
                ["slots"] = new JArray(slots)
            };

            byte[] thumb = TryMakeThumb(channel, foundShape, pageIndex, boxL, boxT, boxW, boxH);
            if (!PptLibStore.TryWriteUser(
                    userId,
                    category,
                    meta,
                    fragment,
                    thumb,
                    overwrite,
                    out bool overwritten,
                    out string writeError))
            {
                errorResult = new ToolResult { Success = false, Error = writeError };
                return false;
            }

            data = new Dictionary<string, object>
            {
                ["id"] = userId,
                ["category"] = category,
                ["has_thumb"] = thumb != null && thumb.Length > 0,
                ["slots"] = slots,
                ["overwritten"] = overwritten
            };
            return true;
        }

        private static bool TryReadExtractTree(
            IOperationChannel channel,
            string slideId,
            int comId,
            out PptHtmlShapeNode node,
            out string error)
        {
            node = null;
            error = null;
            if (channel is PptChannel ppt)
            {
                if (ppt == null || !ppt.TryGetLivePresentation(out PowerPoint.Presentation presentation))
                {
                    error = "渠道对应的演示文稿已关闭";
                    return false;
                }

                return PptHtmlPowerPointReader.TryReadExtractTree(presentation, slideId, comId, out node, out error);
            }

            if (channel is WppChannel wpp)
            {
                if (wpp == null || !wpp.TryGetLivePresentation(out object presentation))
                {
                    error = "渠道对应的演示文稿已关闭";
                    return false;
                }

                return PptHtmlWppReader.TryReadExtractTree(presentation, slideId, comId, out node, out error);
            }

            error = "unsupported: 当前渠道不是 ppt/wpp";
            return false;
        }

        private static void StripShapeIds(PptHtmlShapeNode node)
        {
            if (node == null)
            {
                return;
            }

            node.ShapeId = null;
            if (node.Children == null)
            {
                return;
            }

            foreach (PptHtmlShapeNode child in node.Children)
            {
                StripShapeIds(child);
            }
        }

        private static void TryReadComSlideBox(
            IOperationChannel channel,
            string slideId,
            object foundShape,
            out double left,
            out double top,
            out double width,
            out double height)
        {
            left = top = 0;
            width = height = 100;
            if (foundShape == null)
            {
                return;
            }

            try
            {
                if (foundShape is PowerPoint.Shape pptShape && channel is PptChannel ppt
                    && ppt.TryGetLivePresentation(out PowerPoint.Presentation presentation)
                    && int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                {
                    float sw = presentation.PageSetup.SlideWidth;
                    float sh = presentation.PageSetup.SlideHeight;
                    if (sw > 0 && sh > 0)
                    {
                        left = pptShape.Left / sw * 100.0;
                        top = pptShape.Top / sh * 100.0;
                        width = pptShape.Width / sw * 100.0;
                        height = pptShape.Height / sh * 100.0;
                    }

                    return;
                }

                if (channel is WppChannel wpp && wpp.TryGetLivePresentation(out object wppPres))
                {
                    object pageSetup = WppCom.GetProperty(wppPres, "PageSetup");
                    double sw = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideWidth"));
                    double sh = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideHeight"));
                    if (sw > 0 && sh > 0)
                    {
                        left = Convert.ToDouble(WppCom.GetProperty(foundShape, "Left")) / sw * 100.0;
                        top = Convert.ToDouble(WppCom.GetProperty(foundShape, "Top")) / sh * 100.0;
                        width = Convert.ToDouble(WppCom.GetProperty(foundShape, "Width")) / sw * 100.0;
                        height = Convert.ToDouble(WppCom.GetProperty(foundShape, "Height")) / sh * 100.0;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryCollectIds(
            IOperationChannel channel,
            string slideId,
            int comId,
            out HashSet<int> ids,
            out object foundShape,
            out int pageIndex,
            out string error)
        {
            ids = new HashSet<int>();
            foundShape = null;
            pageIndex = 0;
            error = null;
            if (channel is PptChannel ppt)
            {
                return TryCollectPowerPoint(ppt, slideId, comId, ids, out foundShape, out pageIndex, out error);
            }

            if (channel is WppChannel wpp)
            {
                return TryCollectWpp(wpp, slideId, comId, ids, out foundShape, out pageIndex, out error);
            }

            error = "unsupported: 当前渠道不是 ppt/wpp";
            return false;
        }

        private static bool TryCollectPowerPoint(
            PptChannel ppt,
            string slideId,
            int comId,
            HashSet<int> ids,
            out object foundShape,
            out int pageIndex,
            out string error)
        {
            foundShape = null;
            pageIndex = 0;
            error = null;
            if (ppt == null || !ppt.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (!int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int slideIdInt))
            {
                error = "非法 slide_id: " + slideId;
                return false;
            }

            PowerPoint.Slide slide = null;
            try
            {
                for (int i = 1; i <= presentation.Slides.Count; i++)
                {
                    PowerPoint.Slide candidate = presentation.Slides[i];
                    if (candidate.SlideID == slideIdInt)
                    {
                        slide = candidate;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                error = "查找幻灯片失败: " + ex.Message;
                return false;
            }

            if (slide == null)
            {
                error = "幻灯片不存在: slide_id=" + slideId;
                return false;
            }

            PowerPoint.Shape shape = FindPowerPointShape(slide.Shapes, comId);
            if (shape == null)
            {
                error = "页内找不到 ShapeId（slide_id=" + slideId + "）";
                return false;
            }

            foundShape = shape;
            try
            {
                pageIndex = slide.SlideIndex;
            }
            catch (Exception)
            {
            }

            CollectPowerPointIds(shape, ids);
            return true;
        }

        private static PowerPoint.Shape FindPowerPointShape(PowerPoint.Shapes shapes, int id)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                for (int i = 1; i <= shapes.Count; i++)
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
                            PowerPoint.Shape found = FindPowerPointInGroup(shape, id);
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

        private static PowerPoint.Shape FindPowerPointInGroup(PowerPoint.Shape group, int id)
        {
            try
            {
                for (int i = 1; i <= group.GroupItems.Count; i++)
                {
                    PowerPoint.Shape child = group.GroupItems[i];
                    if (child.Id == id)
                    {
                        return child;
                    }

                    if ((int)child.Type == 6)
                    {
                        PowerPoint.Shape nested = FindPowerPointInGroup(child, id);
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

        private static void CollectPowerPointIds(PowerPoint.Shape shape, HashSet<int> ids)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                ids.Add(shape.Id);
                if ((int)shape.Type == 6)
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        CollectPowerPointIds(shape.GroupItems[i], ids);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryCollectWpp(
            WppChannel wpp,
            string slideId,
            int comId,
            HashSet<int> ids,
            out object foundShape,
            out int pageIndex,
            out string error)
        {
            foundShape = null;
            pageIndex = 0;
            error = null;
            if (wpp == null || !wpp.TryGetLivePresentation(out object presentation))
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            object slides = WppCom.GetProperty(presentation, "Slides");
            if (slides == null)
            {
                error = "无法读取 Slides";
                return false;
            }

            if (!int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int slideIdInt))
            {
                error = "非法 slide_id: " + slideId;
                return false;
            }

            object slide = null;
            try
            {
                int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object candidate = WppCom.GetIndexed(slides, i);
                    object raw = candidate == null ? null : WppCom.GetProperty(candidate, "SlideID");
                    if (raw != null && Convert.ToInt32(raw) == slideIdInt)
                    {
                        slide = candidate;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                error = "查找幻灯片失败: " + ex.Message;
                return false;
            }

            if (slide == null)
            {
                error = "幻灯片不存在: slide_id=" + slideId;
                return false;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            object shape = FindWppShape(shapes, comId);
            if (shape == null)
            {
                error = "页内找不到 ShapeId（slide_id=" + slideId + "）";
                return false;
            }

            foundShape = shape;
            try
            {
                pageIndex = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex"));
            }
            catch (Exception)
            {
            }

            CollectWppIds(shape, ids);
            return true;
        }

        private static object FindWppShape(object shapes, int id)
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
                    object raw = shape == null ? null : WppCom.GetProperty(shape, "Id");
                    if (raw != null && Convert.ToInt32(raw) == id)
                    {
                        return shape;
                    }

                    object type = shape == null ? null : WppCom.GetProperty(shape, "Type");
                    if (type != null && Convert.ToInt32(type) == 6)
                    {
                        object found = FindWppInGroup(shape, id);
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

        private static object FindWppInGroup(object group, int id)
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
                    object raw = child == null ? null : WppCom.GetProperty(child, "Id");
                    if (raw != null && Convert.ToInt32(raw) == id)
                    {
                        return child;
                    }

                    object type = child == null ? null : WppCom.GetProperty(child, "Type");
                    if (type != null && Convert.ToInt32(type) == 6)
                    {
                        object nested = FindWppInGroup(child, id);
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

        private static void CollectWppIds(object shape, HashSet<int> ids)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                object raw = WppCom.GetProperty(shape, "Id");
                if (raw != null)
                {
                    ids.Add(Convert.ToInt32(raw));
                }

                object type = WppCom.GetProperty(shape, "Type");
                if (type != null && Convert.ToInt32(type) == 6)
                {
                    object items = WppCom.GetProperty(shape, "GroupItems");
                    if (items == null)
                    {
                        return;
                    }

                    int count = Convert.ToInt32(WppCom.GetProperty(items, "Count"));
                    for (int i = 1; i <= count; i++)
                    {
                        CollectWppIds(WppCom.GetIndexed(items, i), ids);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static byte[] TryMakeThumb(
            IOperationChannel channel,
            object foundShape,
            int pageIndex,
            double boxL,
            double boxT,
            double boxW,
            double boxH)
        {
            byte[] exported = TryExportShape(foundShape);
            if (exported != null)
            {
                return exported;
            }

            if (pageIndex < 1)
            {
                return null;
            }

            if (!PresentationHostAdapter.TryCaptureSlide(channel, pageIndex, out PresentationCaptureResult cap, out _))
            {
                return null;
            }

            if (cap == null || cap.Image == null || cap.Image.Bytes == null)
            {
                return null;
            }

            try
            {
                using (var stream = new MemoryStream(cap.Image.Bytes, writable: false))
                using (var src = new Bitmap(stream))
                {
                    int x = Math.Max(0, (int)Math.Floor(src.Width * boxL / 100.0));
                    int y = Math.Max(0, (int)Math.Floor(src.Height * boxT / 100.0));
                    int w = Math.Max(1, (int)Math.Ceiling(src.Width * boxW / 100.0));
                    int h = Math.Max(1, (int)Math.Ceiling(src.Height * boxH / 100.0));
                    if (x + w > src.Width)
                    {
                        w = src.Width - x;
                    }

                    if (y + h > src.Height)
                    {
                        h = src.Height - y;
                    }

                    if (w <= 0 || h <= 0)
                    {
                        return null;
                    }

                    using (var crop = src.Clone(new Rectangle(x, y, w, h), src.PixelFormat))
                    {
                        return ImageCaptureCompressor.CompressThumb(crop).Bytes;
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static byte[] TryExportShape(object shape)
        {
            if (shape == null)
            {
                return null;
            }

            string path = Path.Combine(Path.GetTempPath(), "ew-ppt-lib-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                if (shape is PowerPoint.Shape pptShape)
                {
                    pptShape.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 640, 480);
                }
                else
                {
                    WppCom.Invoke(shape, "Export", path, 2, 640, 480);
                }

                if (!File.Exists(path))
                {
                    return null;
                }

                using (var bitmap = new Bitmap(path))
                {
                    return ImageCaptureCompressor.CompressThumb(bitmap).Bytes;
                }
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
