using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class PptSlideWppManager
    {
        public static bool TryManage(
            object dest,
            object source,
            string destChannelId,
            string sourceChannelId,
            PresentationManageSlideRequest request,
            out PresentationManageSlideResult result,
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
                error = "须提供 action（add|duplicate|delete|move|reorder）";
                return false;
            }

            string action = request.Action.Trim().ToLowerInvariant();
            var created = new List<CreatedSlideInfo>();
            var deleted = new List<string>();
            object destSlides = null;
            try
            {
                destSlides = WppCom.GetProperty(dest, "Slides");
                if (destSlides == null)
                {
                    error = "无法读取 Slides";
                    return false;
                }

                object sourceSlides = destSlides;
                if (!ReferenceEquals(dest, source))
                {
                    sourceSlides = WppCom.GetProperty(source, "Slides");
                    if (sourceSlides == null)
                    {
                        error = "无法读取源头 Slides";
                        return false;
                    }
                }

                string focusId = "";
                int? focusIndex = null;
                List<ClearedPlaceholderInfo> cleared = null;

                switch (action)
                {
                    case "add":
                        if (!TryAdd(dest, destSlides, request, out focusId, out focusIndex, out cleared, out error))
                        {
                            return false;
                        }

                        break;
                    case "duplicate":
                        if (!TryDuplicate(
                                dest,
                                source,
                                destSlides,
                                sourceSlides,
                                string.Equals(destChannelId, sourceChannelId, StringComparison.Ordinal),
                                request,
                                created,
                                out focusId,
                                out focusIndex,
                                out error))
                        {
                            result = BuildResult(
                                destChannelId, sourceChannelId, "wpp", action,
                                focusId, focusIndex, destSlides, cleared, created, deleted);
                            return false;
                        }

                        break;
                    case "delete":
                        if (!TryDelete(dest, destSlides, request, deleted, out focusId, out error))
                        {
                            result = BuildResult(
                                destChannelId, sourceChannelId, "wpp", action,
                                focusId, focusIndex, destSlides, cleared, created, deleted);
                            return false;
                        }

                        break;
                    case "move":
                        if (!TryMove(destSlides, request, out focusId, out focusIndex, out error))
                        {
                            return false;
                        }

                        break;
                    case "reorder":
                        if (!TryReorder(destSlides, request, out error))
                        {
                            return false;
                        }

                        break;
                    default:
                        error = "action 须为 add|duplicate|delete|move|reorder";
                        return false;
                }

                result = BuildResult(
                    destChannelId, sourceChannelId, "wpp", action,
                    focusId, focusIndex, destSlides, cleared, created, deleted);
                return true;
            }
            catch (Exception ex)
            {
                error = "管理幻灯片失败: " + ex.Message;
                result = destSlides == null
                    ? null
                    : BuildResult(
                        destChannelId, sourceChannelId, "wpp", action,
                        "", null, destSlides, null, created, deleted);
                return false;
            }
        }

        private static bool TryAdd(
            object presentation,
            object slides,
            PresentationManageSlideRequest request,
            out string focusId,
            out int? focusIndex,
            out List<ClearedPlaceholderInfo> cleared,
            out string error)
        {
            focusId = "";
            focusIndex = null;
            cleared = new List<ClearedPlaceholderInfo>();
            error = null;

            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            int toIndex = request.ToIndex ?? (count + 1);
            if (toIndex < 1 || toIndex > count + 1)
            {
                error = "to_index 越界（add 允许 1.." + (count + 1) + "）";
                return false;
            }

            if (!TryResolveCustomLayout(presentation, request.Layout, out object layout, out error))
            {
                return false;
            }

            object created = WppCom.Invoke(slides, "AddSlide", toIndex, layout);
            if (created == null)
            {
                error = "新建幻灯片失败";
                return false;
            }

            focusId = Convert.ToString(WppCom.GetProperty(created, "SlideID")) ?? "";
            focusIndex = TryGetIndex(created);
            ClearEmptyPlaceholders(created, cleared);
            return true;
        }

        private const int MsoPlaceholder = 14;

        private static void ClearEmptyPlaceholders(object slide, List<ClearedPlaceholderInfo> cleared)
        {
            if (slide == null)
            {
                return;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            if (shapes == null)
            {
                return;
            }

            var toDelete = new List<object>();
            int count;
            try
            {
                count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
            }
            catch (Exception)
            {
                return;
            }

            for (int i = 1; i <= count; i++)
            {
                object shape = WppCom.GetIndexed(shapes, i);
                try
                {
                    object typeObj = WppCom.GetProperty(shape, "Type");
                    if (typeObj == null || Convert.ToInt32(typeObj) != MsoPlaceholder)
                    {
                        continue;
                    }

                    object pf = WppCom.GetProperty(shape, "PlaceholderFormat");
                    object phObj = pf == null ? null : WppCom.GetProperty(pf, "Type");
                    if (phObj == null)
                    {
                        continue;
                    }

                    int ph = Convert.ToInt32(phObj);
                    if (!PptEmptyPlaceholderClear.IsClearableType(ph))
                    {
                        continue;
                    }

                    string text = "";
                    object tf = WppCom.GetProperty(shape, "TextFrame");
                    object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                    if (tr != null)
                    {
                        text = Convert.ToString(WppCom.GetProperty(tr, "Text")) ?? "";
                    }

                    if (!PptEmptyPlaceholderClear.IsEmptyText(text))
                    {
                        continue;
                    }

                    int comId = 0;
                    object idObj = WppCom.GetProperty(shape, "Id");
                    if (idObj != null)
                    {
                        comId = Convert.ToInt32(idObj);
                    }

                    PptEmptyPlaceholderClear.Remember(cleared, ph, comId);
                    toDelete.Add(shape);
                }
                catch (Exception)
                {
                }
            }

            foreach (object shape in toDelete)
            {
                try
                {
                    WppCom.Invoke(shape, "Delete");
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool TryDuplicate(
            object dest,
            object sourcePres,
            object destSlides,
            object sourceSlides,
            bool sameDoc,
            PresentationManageSlideRequest request,
            List<CreatedSlideInfo> created,
            out string focusId,
            out int? focusIndex,
            out string error)
        {
            focusId = "";
            focusIndex = null;
            error = null;

            if (!request.TryResolveSlideSequence(out string[] sequence, out error))
            {
                return false;
            }

            if (!AreSameApplication(dest, sourcePres, out error))
            {
                return false;
            }

            var sourceList = new List<object>(sequence.Length);
            foreach (string id in sequence)
            {
                if (!TryFindSlideById(sourceSlides, id, out object slide, out error))
                {
                    return false;
                }

                sourceList.Add(slide);
            }

            int destCount = Convert.ToInt32(WppCom.GetProperty(destSlides, "Count"));
            int toIndex = request.ToIndex ?? (destCount + 1);
            if (toIndex < 1 || toIndex > destCount + 1)
            {
                error = "to_index 越界（duplicate 允许 1.." + (destCount + 1) + "）";
                return false;
            }

            for (int i = 0; i < sourceList.Count; i++)
            {
                object createdSlide;
                if (sameDoc)
                {
                    if (!TryDuplicateOne(sourceList[i], toIndex, out createdSlide, out error))
                    {
                        return false;
                    }
                }
                else if (!TryCopyFromOne(sourceList[i], dest, destSlides, toIndex, out createdSlide, out error))
                {
                    return false;
                }

                string newId = Convert.ToString(WppCom.GetProperty(createdSlide, "SlideID")) ?? "";
                int newIndex = TryGetIndex(createdSlide);
                created.Add(new CreatedSlideInfo
                {
                    SourceSlideId = sequence[i],
                    SlideId = newId,
                    Index = newIndex
                });
                focusId = newId;
                focusIndex = newIndex;
                toIndex++;
            }

            return true;
        }

        private static bool TryDuplicateOne(
            object source,
            int toIndex,
            out object created,
            out string error)
        {
            created = null;
            error = null;
            object dup = WppCom.Invoke(source, "Duplicate");
            created = ExtractFirstSlide(dup) ?? dup;
            if (created == null)
            {
                error = "复制幻灯片失败";
                return false;
            }

            if (TryGetIndex(created) != toIndex)
            {
                WppCom.Invoke(created, "MoveTo", toIndex);
            }

            return true;
        }

        private static bool TryCopyFromOne(
            object source,
            object dest,
            object destSlides,
            int toIndex,
            out object created,
            out string error)
        {
            created = null;
            error = null;
            object app = null;
            object prevAlerts = null;
            try
            {
                app = WppCom.GetProperty(dest, "Application");
                if (app != null)
                {
                    prevAlerts = WppCom.GetProperty(app, "DisplayAlerts");
                    WppCom.TrySetProperty(app, "DisplayAlerts", false);
                }

                WppCom.Invoke(source, "Copy");
                object pasted = WppCom.Invoke(destSlides, "Paste", toIndex);
                created = ExtractFirstSlide(pasted) ?? pasted;
                if (created == null)
                {
                    error = "跨文档粘贴幻灯片失败";
                    return false;
                }

                ApplyKeepSourceFormatting(source, created);
                if (TryGetIndex(created) != toIndex)
                {
                    WppCom.Invoke(created, "MoveTo", toIndex);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "跨文档复制幻灯片失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (app != null && prevAlerts != null)
                {
                    try
                    {
                        WppCom.TrySetProperty(app, "DisplayAlerts", prevAlerts);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static void ApplyKeepSourceFormatting(object source, object dest)
        {
            try
            {
                object design = WppCom.GetProperty(source, "Design");
                if (design != null)
                {
                    WppCom.TrySetProperty(dest, "Design", design);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object follow = WppCom.GetProperty(source, "FollowMasterBackground");
                if (follow != null)
                {
                    WppCom.TrySetProperty(dest, "FollowMasterBackground", follow);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object scheme = WppCom.GetProperty(source, "ColorScheme");
                if (scheme != null)
                {
                    WppCom.TrySetProperty(dest, "ColorScheme", scheme);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool AreSameApplication(object dest, object source, out string error)
        {
            error = null;
            try
            {
                object destApp = WppCom.GetProperty(dest, "Application");
                object sourceApp = WppCom.GetProperty(source, "Application");
                object destHwnd = destApp == null ? null : WppCom.GetProperty(destApp, "HWND");
                object sourceHwnd = sourceApp == null ? null : WppCom.GetProperty(sourceApp, "HWND");
                if (destHwnd != null && sourceHwnd != null
                    && Convert.ToInt32(destHwnd) != Convert.ToInt32(sourceHwnd))
                {
                    error = "源头与目标不在同一演示文稿应用内";
                    return false;
                }
            }
            catch (Exception)
            {
            }

            return true;
        }

        private static bool TryDelete(
            object presentation,
            object slides,
            PresentationManageSlideRequest request,
            List<string> deleted,
            out string focusId,
            out string error)
        {
            focusId = "";
            error = null;

            if (!request.Confirm)
            {
                error = "删除须 confirm=true";
                return false;
            }

            if (!request.TryResolveSlideSequence(out string[] sequence, out error))
            {
                return false;
            }

            if (PresentationManageSlideRequest.SequenceHasDuplicate(sequence, out string dupId))
            {
                error = "slide_ids 含重复 slide_id: " + dupId;
                return false;
            }

            var slideList = new List<object>(sequence.Length);
            foreach (string id in sequence)
            {
                if (!TryFindSlideById(slides, id, out object slide, out error))
                {
                    return false;
                }

                slideList.Add(slide);
            }

            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            if (sequence.Length >= count)
            {
                error = "不能删除演示文稿中唯一的幻灯片";
                return false;
            }

            object app = null;
            object prevAlerts = null;
            try
            {
                app = WppCom.GetProperty(presentation, "Application");
                if (app != null)
                {
                    prevAlerts = WppCom.GetProperty(app, "DisplayAlerts");
                    WppCom.TrySetProperty(app, "DisplayAlerts", false);
                }

                for (int i = 0; i < slideList.Count; i++)
                {
                    WppCom.Invoke(slideList[i], "Delete");
                    deleted.Add(sequence[i]);
                    focusId = sequence[i];
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "删除幻灯片失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (app != null && prevAlerts != null)
                {
                    try
                    {
                        WppCom.TrySetProperty(app, "DisplayAlerts", prevAlerts);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool TryMove(
            object slides,
            PresentationManageSlideRequest request,
            out string focusId,
            out int? focusIndex,
            out string error)
        {
            focusId = "";
            focusIndex = null;
            error = null;

            if (string.IsNullOrWhiteSpace(request.SlideId))
            {
                error = "move 须提供 slide_id";
                return false;
            }

            if (!request.ToIndex.HasValue)
            {
                error = "move 须提供 to_index";
                return false;
            }

            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            int toIndex = request.ToIndex.Value;
            if (toIndex < 1 || toIndex > count)
            {
                error = "to_index 越界（move 允许 1.." + count + "）";
                return false;
            }

            if (!TryFindSlideById(slides, request.SlideId.Trim(), out object slide, out error))
            {
                return false;
            }

            focusId = Convert.ToString(WppCom.GetProperty(slide, "SlideID")) ?? request.SlideId.Trim();
            int current = TryGetIndex(slide);
            if (current == toIndex)
            {
                focusIndex = toIndex;
                return true;
            }

            WppCom.Invoke(slide, "MoveTo", toIndex);
            focusIndex = TryGetIndex(slide);
            return true;
        }

        private static bool TryReorder(
            object slides,
            PresentationManageSlideRequest request,
            out string error)
        {
            error = null;
            if (request.Order == null || request.Order.Length == 0)
            {
                error = "reorder 须提供 order（全部 slide_id 的排列）";
                return false;
            }

            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            if (request.Order.Length != count)
            {
                error = "order 长度须等于当前页数 " + count;
                return false;
            }

            var currentIds = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 1; i <= count; i++)
            {
                object s = WppCom.GetIndexed(slides, i);
                currentIds.Add(Convert.ToString(WppCom.GetProperty(s, "SlideID")) ?? "");
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string raw in request.Order)
            {
                string id = (raw ?? "").Trim();
                if (string.IsNullOrEmpty(id))
                {
                    error = "order 含空 slide_id";
                    return false;
                }

                if (!seen.Add(id))
                {
                    error = "order 含重复 slide_id: " + id;
                    return false;
                }

                if (!currentIds.Contains(id))
                {
                    error = "order 含未知 slide_id: " + id;
                    return false;
                }
            }

            if (seen.Count != currentIds.Count)
            {
                error = "order 须覆盖当前全部 slide_id";
                return false;
            }

            for (int targetPos = 1; targetPos <= count; targetPos++)
            {
                string wantId = request.Order[targetPos - 1].Trim();
                if (!TryFindSlideById(slides, wantId, out object slide, out error))
                {
                    return false;
                }

                if (TryGetIndex(slide) != targetPos)
                {
                    WppCom.Invoke(slide, "MoveTo", targetPos);
                }
            }

            return true;
        }

        private static bool TryResolveCustomLayout(
            object presentation,
            string layoutName,
            out object layout,
            out string error)
        {
            layout = null;
            error = null;

            object master;
            object layouts;
            try
            {
                master = WppCom.GetProperty(presentation, "SlideMaster");
                layouts = WppCom.GetProperty(master, "CustomLayouts");
            }
            catch (Exception ex)
            {
                error = "读取版式失败: " + ex.Message;
                return false;
            }

            if (layouts == null)
            {
                error = "演示文稿无可用版式";
                return false;
            }

            int layoutCount = Convert.ToInt32(WppCom.GetProperty(layouts, "Count"));
            if (layoutCount < 1)
            {
                error = "演示文稿无可用版式";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(layoutName))
            {
                string want = layoutName.Trim();
                for (int i = 1; i <= layoutCount; i++)
                {
                    object candidate = WppCom.GetIndexed(layouts, i);
                    string name = Convert.ToString(WppCom.GetProperty(candidate, "Name")) ?? "";
                    if (string.Equals(name, want, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(name, want, StringComparison.Ordinal))
                    {
                        layout = candidate;
                        return true;
                    }
                }

                error = "版式不存在: " + want;
                return false;
            }

            for (int i = 1; i <= layoutCount; i++)
            {
                object candidate = WppCom.GetIndexed(layouts, i);
                string name = Convert.ToString(WppCom.GetProperty(candidate, "Name")) ?? "";
                if (name.IndexOf("空白", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Blank", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    layout = candidate;
                    return true;
                }
            }

            layout = WppCom.GetIndexed(layouts, 1);
            return layout != null;
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

        private static object ExtractFirstSlide(object dup)
        {
            if (dup == null)
            {
                return null;
            }

            try
            {
                object countObj = WppCom.GetProperty(dup, "Count");
                if (countObj != null && Convert.ToInt32(countObj) >= 1)
                {
                    return WppCom.GetIndexed(dup, 1);
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static int TryGetIndex(object slide)
        {
            try
            {
                object index = WppCom.GetProperty(slide, "SlideIndex");
                if (index != null)
                {
                    return Convert.ToInt32(index);
                }
            }
            catch (Exception)
            {
            }

            return 0;
        }

        private static PresentationManageSlideResult BuildResult(
            string channelId,
            string sourceChannelId,
            string kind,
            string action,
            string focusId,
            int? focusIndex,
            object slides,
            List<ClearedPlaceholderInfo> cleared,
            List<CreatedSlideInfo> created,
            List<string> deleted)
        {
            var list = new List<PresentationSlideInfo>();
            int count = 0;
            try
            {
                count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
                int limit = Math.Min(count, PresentationContentResult.MaxSlides);
                for (int i = 1; i <= limit; i++)
                {
                    object s = WppCom.GetIndexed(slides, i);
                    list.Add(ReadLight(s, i));
                }
            }
            catch (Exception)
            {
            }

            return new PresentationManageSlideResult
            {
                ChannelId = channelId ?? "",
                SourceChannelId = sourceChannelId ?? "",
                Kind = kind,
                Action = action,
                FocusSlideId = focusId ?? "",
                FocusIndex = focusIndex,
                SlideCount = count,
                Slides = list,
                ClearedPlaceholders = cleared,
                Created = created,
                Deleted = deleted
            };
        }

        private static PresentationSlideInfo ReadLight(object slide, int fallbackIndex)
        {
            var info = new PresentationSlideInfo
            {
                Index = fallbackIndex,
                SlideId = "",
                Title = "",
                Layout = "",
                Hidden = false
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
                info.SlideId = Convert.ToString(WppCom.GetProperty(slide, "SlideID")) ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                object transition = WppCom.GetProperty(slide, "SlideShowTransition");
                object hidden = WppCom.GetProperty(transition, "Hidden");
                info.Hidden = IsTruthy(hidden);
            }
            catch (Exception)
            {
            }

            try
            {
                object layout = WppCom.GetProperty(slide, "CustomLayout");
                info.Layout = Convert.ToString(WppCom.GetProperty(layout, "Name")) ?? "";
            }
            catch (Exception)
            {
            }

            return info;
        }

        private static bool IsTruthy(object value)
        {
            if (value == null)
            {
                return false;
            }

            if (value is bool b)
            {
                return b;
            }

            try
            {
                int n = Convert.ToInt32(value);
                return n == -1 || n == 1;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
