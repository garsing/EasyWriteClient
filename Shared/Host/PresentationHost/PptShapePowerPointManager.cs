using System;
using System.Collections.Generic;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    internal static class PptShapePowerPointManager
    {
        private const int MsoGroup = 6;

        public static bool TryManage(
            PowerPoint.Presentation dest,
            PowerPoint.Presentation source,
            string destChannelId,
            string sourceChannelId,
            PresentationManageShapeRequest request,
            out PresentationManageShapeResult result,
            out string error)
        {
            result = null;
            error = null;
            if (dest == null)
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (source == null)
            {
                source = dest;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "须提供 action（delete、group 或 duplicate_group）";
                return false;
            }

            string action = request.Action.Trim().ToLowerInvariant();
            switch (action)
            {
                case "delete":
                    return TryDelete(dest, destChannelId, request, out result, out error);
                case "group":
                    return TryGroup(dest, destChannelId, request, out result, out error);
                case "duplicate_group":
                    return TryDuplicateGroup(
                        dest,
                        source,
                        destChannelId,
                        sourceChannelId,
                        request,
                        out result,
                        out error);
                default:
                    error = "action 须为 delete、group 或 duplicate_group";
                    return false;
            }
        }

        private static bool TryDelete(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationManageShapeRequest request,
            out PresentationManageShapeResult result,
            out string error)
        {
            result = null;
            error = null;
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

            result = BaseResult(channelId, "delete", request.SlideId.Trim());
            result.DeletedCount = deleted.Count;
            result.DeletedShapeIds = deleted;
            return true;
        }

        private static bool TryGroup(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationManageShapeRequest request,
            out PresentationManageShapeResult result,
            out string error)
        {
            result = null;
            error = null;
            if (string.IsNullOrWhiteSpace(request.SlideId))
            {
                error = "须提供 slide_id";
                return false;
            }

            if (request.ShapeIds == null || request.ShapeIds.Length < 2)
            {
                error = "group 须提供至少 2 个 shape_ids";
                return false;
            }

            if (!TryFindSlideById(presentation, request.SlideId.Trim(), out PowerPoint.Slide slide, out error))
            {
                return false;
            }

            var members = new List<PowerPoint.Shape>();
            var names = new List<object>();
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
                    error = "group 的 shape_ids 不能重复";
                    return false;
                }

                PowerPoint.Shape top = FindTopLevelShapeById(slide.Shapes, comId);
                if (top == null)
                {
                    if (FindShapeById(slide.Shapes, comId) != null)
                    {
                        error = "只能把本页顶层形状组成组，ShapeId 已在组内: " + (raw ?? "");
                    }
                    else
                    {
                        error = "页内找不到顶层 ShapeId=" + (raw ?? "") + "（slide_id=" + request.SlideId + "）";
                    }

                    return false;
                }

                members.Add(top);
                names.Add(top.Name);
            }

            if (members.Count < 2)
            {
                error = "group 须提供至少 2 个 shape_ids";
                return false;
            }

            PowerPoint.Shape group = null;
            PowerPoint.PpAlertLevel prevAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
            bool alertsSet = false;
            try
            {
                PowerPoint.Application app = presentation.Application;
                if (app != null)
                {
                    prevAlerts = app.DisplayAlerts;
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                    alertsSet = true;
                }

                group = slide.Shapes.Range(names.ToArray()).Group();
                if (group == null)
                {
                    error = "COM Group 失败";
                    return false;
                }

                result = BaseResult(channelId, "group", request.SlideId.Trim());
                result.GroupShapeId = PptShapeId.FormatShape(request.SlideId.Trim(), group.Id);
                return true;
            }
            catch (Exception ex)
            {
                error = "组成组失败: " + ex.Message;
                if (group != null)
                {
                    try
                    {
                        group.Delete();
                    }
                    catch (Exception)
                    {
                    }

                    error += "；页可能已成组";
                }

                return false;
            }
            finally
            {
                if (alertsSet)
                {
                    try
                    {
                        presentation.Application.DisplayAlerts = prevAlerts;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool TryDuplicateGroup(
            PowerPoint.Presentation dest,
            PowerPoint.Presentation source,
            string destChannelId,
            string sourceChannelId,
            PresentationManageShapeRequest request,
            out PresentationManageShapeResult result,
            out string error)
        {
            result = null;
            error = null;
            if (string.IsNullOrWhiteSpace(request.SlideId))
            {
                error = "须提供 slide_id（源页）";
                return false;
            }

            if (string.IsNullOrWhiteSpace(request.ShapeId))
            {
                error = "须提供 shape_id（源组）";
                return false;
            }

            if (!request.LeftPct.HasValue || !request.TopPct.HasValue)
            {
                error = "left/top 须为相对目标幻灯片的百分比";
                return false;
            }

            if (!SameApplication(dest, source))
            {
                error = "源头与目标须在同一 PowerPoint 进程内";
                return false;
            }

            string sourceSlideId = request.SlideId.Trim();
            string destSlideId = string.IsNullOrWhiteSpace(request.ToSlideId)
                ? sourceSlideId
                : request.ToSlideId.Trim();

            if (!TryFindSlideById(source, sourceSlideId, out PowerPoint.Slide sourceSlide, out error))
            {
                return false;
            }

            if (!TryFindSlideById(dest, destSlideId, out PowerPoint.Slide destSlide, out error))
            {
                return false;
            }

            if (!PptShapeId.TryParseShapeComId(request.ShapeId, out int comId))
            {
                error = "非法 ShapeId: " + request.ShapeId;
                return false;
            }

            PowerPoint.Shape sourceShape = FindShapeById(sourceSlide.Shapes, comId);
            if (sourceShape == null)
            {
                error = "页内找不到 ShapeId=" + request.ShapeId + "（slide_id=" + sourceSlideId + "）";
                return false;
            }

            if (!IsGroupShape(sourceShape))
            {
                error = "只能拷贝 group，请先 group";
                return false;
            }

            float destWidth;
            float destHeight;
            try
            {
                destWidth = dest.PageSetup.SlideWidth;
                destHeight = dest.PageSetup.SlideHeight;
            }
            catch (Exception ex)
            {
                error = "无法读取目标页尺寸: " + ex.Message;
                return false;
            }

            float left = (float)(request.LeftPct.Value / 100.0 * destWidth);
            float top = (float)(request.TopPct.Value / 100.0 * destHeight);

            PowerPoint.Shape pasted = null;
            PowerPoint.PpAlertLevel prevAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
            bool alertsSet = false;
            try
            {
                PowerPoint.Application app = dest.Application;
                if (app != null)
                {
                    prevAlerts = app.DisplayAlerts;
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                    alertsSet = true;
                }

                sourceShape.Copy();
                PowerPoint.ShapeRange range = destSlide.Shapes.Paste();
                if (range == null || range.Count < 1)
                {
                    error = "粘贴组失败";
                    return false;
                }

                pasted = range[1];
                if (!IsGroupShape(pasted))
                {
                    try
                    {
                        pasted.Delete();
                    }
                    catch (Exception)
                    {
                    }

                    pasted = null;
                    error = "粘贴结果不是 group";
                    return false;
                }

                pasted.Left = left;
                pasted.Top = top;

                result = BaseResult(destChannelId, "duplicate_group", destSlideId);
                result.GroupShapeId = PptShapeId.FormatShape(destSlideId, pasted.Id);
                result.SourceSlideId = sourceSlideId;
                result.SourceChannelId = sourceChannelId ?? "";
                result.SourceShapeId = PptShapeId.FormatShape(sourceSlideId, comId);
                return true;
            }
            catch (Exception ex)
            {
                error = "拷组失败: " + ex.Message;
                if (pasted != null)
                {
                    try
                    {
                        pasted.Delete();
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            }
            finally
            {
                if (alertsSet)
                {
                    try
                    {
                        dest.Application.DisplayAlerts = prevAlerts;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static PresentationManageShapeResult BaseResult(string channelId, string action, string slideId)
        {
            return new PresentationManageShapeResult
            {
                ChannelId = channelId ?? "",
                Kind = "ppt",
                Action = action,
                SlideId = slideId ?? "",
                DeletedCount = 0,
                DeletedShapeIds = new List<string>(),
                Warnings = new List<string>()
            };
        }

        private static bool SameApplication(PowerPoint.Presentation dest, PowerPoint.Presentation source)
        {
            if (dest == null || source == null)
            {
                return false;
            }

            if (ReferenceEquals(dest, source))
            {
                return true;
            }

            try
            {
                return dest.Application.HWND == source.Application.HWND;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsGroupShape(PowerPoint.Shape shape)
        {
            try
            {
                return shape != null && (int)shape.Type == MsoGroup;
            }
            catch (Exception)
            {
                return false;
            }
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

        private static PowerPoint.Shape FindTopLevelShapeById(PowerPoint.Shapes shapes, int id)
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

                        if ((int)shape.Type == MsoGroup)
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

                    if ((int)child.Type == MsoGroup)
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
