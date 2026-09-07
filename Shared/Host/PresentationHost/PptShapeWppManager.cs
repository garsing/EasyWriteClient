using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class PptShapeWppManager
    {
        private const int MsoGroup = 6;

        public static bool TryManage(
            object dest,
            object source,
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
            object presentation,
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

            result = BaseResult(channelId, "delete", request.SlideId.Trim());
            result.DeletedCount = deleted.Count;
            result.DeletedShapeIds = deleted;
            return true;
        }

        private static bool TryGroup(
            object presentation,
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

            var members = new List<object>();
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

                object top = FindTopLevelShapeById(shapes, comId);
                if (top == null)
                {
                    if (FindShapeById(shapes, comId) != null)
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
                names.Add(WppCom.GetProperty(top, "Name"));
            }

            if (members.Count < 2)
            {
                error = "group 须提供至少 2 个 shape_ids";
                return false;
            }

            object group = null;
            object prevAlerts = null;
            bool alertsSet = false;
            try
            {
                alertsSet = TrySilenceAlerts(presentation, out prevAlerts);
                object range = WppCom.Invoke(shapes, "Range", new object[] { names.ToArray() });
                group = WppCom.Invoke(range, "Group");
                if (group == null)
                {
                    error = "WPP Group 失败";
                    return false;
                }

                object rawId = WppCom.GetProperty(group, "Id");
                if (rawId == null)
                {
                    error = "WPP Group 后读不到新组 Id";
                    try
                    {
                        WppCom.Invoke(group, "Delete");
                    }
                    catch (Exception)
                    {
                    }

                    return false;
                }

                result = BaseResult(channelId, "group", request.SlideId.Trim());
                result.GroupShapeId = PptShapeId.FormatShape(request.SlideId.Trim(), Convert.ToInt32(rawId));
                return true;
            }
            catch (Exception ex)
            {
                error = "组成组失败: " + ex.Message;
                if (group != null)
                {
                    try
                    {
                        WppCom.Invoke(group, "Delete");
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
                    RestoreAlerts(presentation, prevAlerts);
                }
            }
        }

        private static bool TryDuplicateGroup(
            object dest,
            object source,
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
                error = "源头与目标须在同一 WPP 进程内";
                return false;
            }

            string sourceSlideId = request.SlideId.Trim();
            string destSlideId = string.IsNullOrWhiteSpace(request.ToSlideId)
                ? sourceSlideId
                : request.ToSlideId.Trim();

            object sourceSlides = WppCom.GetProperty(source, "Slides");
            object destSlides = WppCom.GetProperty(dest, "Slides");
            if (sourceSlides == null || destSlides == null)
            {
                error = "无法读取 Slides";
                return false;
            }

            if (!TryFindSlideById(sourceSlides, sourceSlideId, out object sourceSlide, out error))
            {
                return false;
            }

            if (!TryFindSlideById(destSlides, destSlideId, out object destSlide, out error))
            {
                return false;
            }

            if (!PptShapeId.TryParseShapeComId(request.ShapeId, out int comId))
            {
                error = "非法 ShapeId: " + request.ShapeId;
                return false;
            }

            object sourceShapes = WppCom.GetProperty(sourceSlide, "Shapes");
            object destShapes = WppCom.GetProperty(destSlide, "Shapes");
            if (sourceShapes == null || destShapes == null)
            {
                error = "无法读取 Shapes";
                return false;
            }

            object sourceShape = FindShapeById(sourceShapes, comId);
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

            double destWidth;
            double destHeight;
            try
            {
                object pageSetup = WppCom.GetProperty(dest, "PageSetup");
                destWidth = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideWidth"));
                destHeight = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideHeight"));
            }
            catch (Exception ex)
            {
                error = "无法读取目标页尺寸: " + ex.Message;
                return false;
            }

            double left = request.LeftPct.Value / 100.0 * destWidth;
            double top = request.TopPct.Value / 100.0 * destHeight;

            object pasted = null;
            object prevAlerts = null;
            bool alertsSet = false;
            try
            {
                alertsSet = TrySilenceAlerts(dest, out prevAlerts);
                WppCom.Invoke(sourceShape, "Copy");
                object pastedRaw = WppCom.Invoke(destShapes, "Paste");
                pasted = FirstPastedShape(pastedRaw, destShapes);
                if (pasted == null)
                {
                    error = "粘贴组失败";
                    return false;
                }

                if (!IsGroupShape(pasted))
                {
                    try
                    {
                        WppCom.Invoke(pasted, "Delete");
                    }
                    catch (Exception)
                    {
                    }

                    pasted = null;
                    error = "粘贴结果不是 group";
                    return false;
                }

                WppCom.TrySetProperty(pasted, "Left", left);
                WppCom.TrySetProperty(pasted, "Top", top);

                object rawId = WppCom.GetProperty(pasted, "Id");
                if (rawId == null)
                {
                    try
                    {
                        WppCom.Invoke(pasted, "Delete");
                    }
                    catch (Exception)
                    {
                    }

                    pasted = null;
                    error = "粘贴后读不到新组 Id";
                    return false;
                }

                result = BaseResult(destChannelId, "duplicate_group", destSlideId);
                result.GroupShapeId = PptShapeId.FormatShape(destSlideId, Convert.ToInt32(rawId));
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
                        WppCom.Invoke(pasted, "Delete");
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
                    RestoreAlerts(dest, prevAlerts);
                }
            }
        }

        private static object FirstPastedShape(object pastedRaw, object destShapes)
        {
            if (pastedRaw != null)
            {
                try
                {
                    object countObj = WppCom.GetProperty(pastedRaw, "Count");
                    if (countObj != null && Convert.ToInt32(countObj) >= 1)
                    {
                        object first = WppCom.GetIndexed(pastedRaw, 1);
                        if (first != null)
                        {
                            return first;
                        }
                    }
                }
                catch (Exception)
                {
                }

                if (WppCom.GetProperty(pastedRaw, "Id") != null)
                {
                    return pastedRaw;
                }
            }

            try
            {
                int count = Convert.ToInt32(WppCom.GetProperty(destShapes, "Count"));
                if (count >= 1)
                {
                    return WppCom.GetIndexed(destShapes, count);
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static bool TrySilenceAlerts(object presentation, out object prev)
        {
            prev = null;
            try
            {
                object app = WppCom.GetProperty(presentation, "Application");
                if (app == null)
                {
                    return false;
                }

                prev = WppCom.GetProperty(app, "DisplayAlerts");
                WppCom.TrySetProperty(app, "DisplayAlerts", false);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void RestoreAlerts(object presentation, object prev)
        {
            try
            {
                object app = WppCom.GetProperty(presentation, "Application");
                if (app != null)
                {
                    WppCom.TrySetProperty(app, "DisplayAlerts", prev);
                }
            }
            catch (Exception)
            {
            }
        }

        private static PresentationManageShapeResult BaseResult(string channelId, string action, string slideId)
        {
            return new PresentationManageShapeResult
            {
                ChannelId = channelId ?? "",
                Kind = "wpp",
                Action = action,
                SlideId = slideId ?? "",
                DeletedCount = 0,
                DeletedShapeIds = new List<string>(),
                Warnings = new List<string>()
            };
        }

        private static bool SameApplication(object dest, object source)
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
                object destApp = WppCom.GetProperty(dest, "Application");
                object sourceApp = WppCom.GetProperty(source, "Application");
                object destHwnd = destApp == null ? null : WppCom.GetProperty(destApp, "HWND");
                object sourceHwnd = sourceApp == null ? null : WppCom.GetProperty(sourceApp, "HWND");
                if (destHwnd == null || sourceHwnd == null)
                {
                    return destApp != null && sourceApp != null;
                }

                return Convert.ToInt32(destHwnd) == Convert.ToInt32(sourceHwnd);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsGroupShape(object shape)
        {
            try
            {
                object type = shape == null ? null : WppCom.GetProperty(shape, "Type");
                return type != null && Convert.ToInt32(type) == MsoGroup;
            }
            catch (Exception)
            {
                return false;
            }
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

        private static object FindTopLevelShapeById(object shapes, int id)
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
                }
            }
            catch (Exception)
            {
            }

            return null;
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
                    if (type != null && Convert.ToInt32(type) == MsoGroup)
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
                    if (type != null && Convert.ToInt32(type) == MsoGroup)
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
