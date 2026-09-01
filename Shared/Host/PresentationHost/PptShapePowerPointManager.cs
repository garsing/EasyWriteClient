using System;
using System.Collections.Generic;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    internal static class PptShapePowerPointManager
    {
        public static bool TryManage(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationManageShapeRequest request,
            out PresentationManageShapeResult result,
            out string error)
        {
            result = null;
            error = null;
            if (presentation == null)
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "须提供 action（delete）";
                return false;
            }

            string action = request.Action.Trim().ToLowerInvariant();
            if (action != "delete")
            {
                error = "action 须为 delete";
                return false;
            }

            if (string.IsNullOrWhiteSpace(request.SlideId))
            {
                error = "须提供 slide_id";
                return false;
            }

            if (request.ShapeIds == null || request.ShapeIds.Length == 0)
            {
                error = "须提供 shape_ids";
                return false;
            }

            if (!TryFindSlideById(presentation, request.SlideId.Trim(), out PowerPoint.Slide slide, out error))
            {
                return false;
            }

            var planned = new List<KeyValuePair<int, string>>();
            var seen = new HashSet<int>();
            foreach (string raw in request.ShapeIds)
            {
                if (!PptShapeId.TryParseShapeComId(raw, out int comId))
                {
                    error = "非法 ShapeId: " + (raw ?? "");
                    return false;
                }

                if (!seen.Add(comId))
                {
                    continue;
                }

                if (FindShapeById(slide.Shapes, comId) == null)
                {
                    error = "页内找不到 ShapeId=" + (raw ?? "") + "（slide_id=" + request.SlideId + "）";
                    return false;
                }

                planned.Add(new KeyValuePair<int, string>(comId, raw ?? ""));
            }

            var deleted = new List<string>();
            try
            {
                foreach (KeyValuePair<int, string> one in planned)
                {
                    PowerPoint.Shape shape = FindShapeById(slide.Shapes, one.Key);
                    if (shape == null)
                    {
                        error = "删除前形状消失: " + one.Value;
                        return false;
                    }

                    shape.Delete();
                    deleted.Add(PptShapeId.FormatShape(request.SlideId.Trim(), one.Key));
                }
            }
            catch (Exception ex)
            {
                error = "删除形状失败: " + ex.Message;
                return false;
            }

            result = new PresentationManageShapeResult
            {
                ChannelId = channelId ?? "",
                Kind = "ppt",
                Action = "delete",
                SlideId = request.SlideId.Trim(),
                DeletedCount = deleted.Count,
                DeletedShapeIds = deleted,
                Warnings = new List<string>()
            };
            return true;
        }

        private static bool TryFindSlideById(
            PowerPoint.Presentation presentation,
            string slideId,
            out PowerPoint.Slide slide,
            out string error)
        {
            slide = null;
            error = null;
            if (!int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idInt))
            {
                error = "非法 slide_id: " + slideId;
                return false;
            }

            try
            {
                for (int i = 1; i <= presentation.Slides.Count; i++)
                {
                    PowerPoint.Slide candidate = presentation.Slides[i];
                    if (candidate.SlideID == idInt)
                    {
                        slide = candidate;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                error = "查找幻灯片失败: " + ex.Message;
                return false;
            }

            error = "幻灯片不存在: slide_id=" + slideId;
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
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PowerPoint.Shape shape = shapes[i];
                    try
                    {
                        if (shape.Id == id)
                        {
                            return shape;
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
    }
}
