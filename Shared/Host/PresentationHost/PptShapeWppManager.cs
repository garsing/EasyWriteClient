using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class PptShapeWppManager
    {
        public static bool TryManage(
            object presentation,
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

            object slides = WppCom.GetProperty(presentation, "Slides");
            if (slides == null)
            {
                error = "无法读取 Slides";
                return false;
            }

            if (!TryFindSlideById(slides, request.SlideId.Trim(), out object slide, out error))
            {
                return false;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            if (shapes == null)
            {
                error = "无法读取 Shapes";
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

                if (FindShapeById(shapes, comId) == null)
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
                    object shape = FindShapeById(shapes, one.Key);
                    if (shape == null)
                    {
                        error = "删除前形状消失: " + one.Value;
                        return false;
                    }

                    WppCom.Invoke(shape, "Delete");
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
                Kind = "wpp",
                Action = "delete",
                SlideId = request.SlideId.Trim(),
                DeletedCount = deleted.Count,
                DeletedShapeIds = deleted,
                Warnings = new List<string>()
            };
            return true;
        }

        private static bool TryFindSlideById(
            object slides,
            string slideId,
            out object slide,
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
                int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object candidate = WppCom.GetIndexed(slides, i);
                    object rawId = WppCom.GetProperty(candidate, "SlideID");
                    if (rawId != null && Convert.ToInt32(rawId) == idInt)
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
                    object raw = WppCom.GetProperty(shape, "Id");
                    if (raw != null && Convert.ToInt32(raw) == id)
                    {
                        return shape;
                    }

                    object type = WppCom.GetProperty(shape, "Type");
                    if (type != null && Convert.ToInt32(type) == 6)
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
                    object raw = child == null ? null : WppCom.GetProperty(child, "Id");
                    if (raw != null && Convert.ToInt32(raw) == id)
                    {
                        return child;
                    }

                    object type = child == null ? null : WppCom.GetProperty(child, "Type");
                    if (type != null && Convert.ToInt32(type) == 6)
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
