using System;
using System.Collections.Generic;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PptSlidePowerPointManager
    {
        public static bool TryManage(
            PowerPoint.Presentation dest,
            PowerPoint.Presentation source,
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
            try
            {
                string focusId = "";
                int? focusIndex = null;
                List<ClearedPlaceholderInfo> cleared = null;

                switch (action)
                {
                    case "add":
                        if (!TryAdd(dest, request, out focusId, out focusIndex, out cleared, out error))
                        {
                            return false;
                        }

                        break;
                    case "duplicate":
                        if (!TryDuplicate(
                                dest,
                                source,
                                string.Equals(destChannelId, sourceChannelId, StringComparison.Ordinal),
                                request,
                                created,
                                out focusId,
                                out focusIndex,
                                out error))
                        {
                            result = BuildResult(
                                destChannelId, sourceChannelId, "ppt", action,
                                focusId, focusIndex, dest, cleared, created, deleted);
                            return false;
                        }

                        break;
                    case "delete":
                        if (!TryDelete(dest, request, deleted, out focusId, out error))
                        {
                            result = BuildResult(
                                destChannelId, sourceChannelId, "ppt", action,
                                focusId, focusIndex, dest, cleared, created, deleted);
                            return false;
                        }

                        break;
                    case "move":
                        if (!TryMove(dest, request, out focusId, out focusIndex, out error))
                        {
                            return false;
                        }

                        break;
                    case "reorder":
                        if (!TryReorder(dest, request, out error))
                        {
                            return false;
                        }

                        break;
                    default:
                        error = "action 须为 add|duplicate|delete|move|reorder";
                        return false;
                }

                result = BuildResult(
                    destChannelId, sourceChannelId, "ppt", action,
                    focusId, focusIndex, dest, cleared, created, deleted);
                return true;
            }
            catch (Exception ex)
            {
                error = "管理幻灯片失败: " + ex.Message;
                result = BuildResult(
                    destChannelId, sourceChannelId, "ppt", action,
                    "", null, dest, null, created, deleted);
                return false;
            }
        }

        private static bool TryAdd(
            PowerPoint.Presentation presentation,
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

            int count = presentation.Slides.Count;
            int toIndex = request.ToIndex ?? (count + 1);
            if (toIndex < 1 || toIndex > count + 1)
            {
                error = "to_index 越界（add 允许 1.." + (count + 1) + "）";
                return false;
            }

            if (!TryResolveCustomLayout(presentation, request.Layout, out PowerPoint.CustomLayout layout, out error))
            {
                return false;
            }

            PowerPoint.Slide created = presentation.Slides.AddSlide(toIndex, layout);
            focusId = Convert.ToString(created.SlideID) ?? "";
            focusIndex = created.SlideIndex;
            ClearEmptyPlaceholders(created, cleared);
            return true;
        }

        private static void ClearEmptyPlaceholders(PowerPoint.Slide slide, List<ClearedPlaceholderInfo> cleared)
        {
            if (slide == null)
            {
                return;
            }

            var toDelete = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                try
                {
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                    {
                        continue;
                    }

                    int ph = Convert.ToInt32(shape.PlaceholderFormat.Type);
                    if (!PptEmptyPlaceholderClear.IsClearableType(ph))
                    {
                        continue;
                    }

                    string text = "";
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        text = shape.TextFrame.TextRange.Text;
                    }

                    if (!PptEmptyPlaceholderClear.IsEmptyText(text))
                    {
                        continue;
                    }

                    PptEmptyPlaceholderClear.Remember(cleared, ph, shape.Id);
                    toDelete.Add(shape);
                }
                catch (Exception)
                {
                }
            }

            foreach (PowerPoint.Shape shape in toDelete)
            {
                try
                {
                    shape.Delete();
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool TryDuplicate(
            PowerPoint.Presentation dest,
            PowerPoint.Presentation source,
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

            if (!AreSameApplication(dest, source, out error))
            {
                return false;
            }

            var sourceSlides = new List<PowerPoint.Slide>(sequence.Length);
            foreach (string id in sequence)
            {
                if (!TryFindSlideById(source, id, out PowerPoint.Slide slide, out error))
                {
                    return false;
                }

                sourceSlides.Add(slide);
            }

            int destCount = dest.Slides.Count;
            int toIndex = request.ToIndex ?? (destCount + 1);
            if (toIndex < 1 || toIndex > destCount + 1)
            {
                error = "to_index 越界（duplicate 允许 1.." + (destCount + 1) + "）";
                return false;
            }

            for (int i = 0; i < sourceSlides.Count; i++)
            {
                PowerPoint.Slide createdSlide;
                if (sameDoc)
                {
                    if (!TryDuplicateOne(sourceSlides[i], toIndex, out createdSlide, out error))
                    {
                        return false;
                    }
                }
                else if (!TryCopyFromOne(sourceSlides[i], dest, toIndex, out createdSlide, out error))
                {
                    return false;
                }

                string newId = Convert.ToString(createdSlide.SlideID) ?? "";
                int newIndex = createdSlide.SlideIndex;
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
            PowerPoint.Slide source,
            int toIndex,
            out PowerPoint.Slide created,
            out string error)
        {
            created = null;
            error = null;
            PowerPoint.SlideRange dup = source.Duplicate();
            if (dup == null || dup.Count < 1)
            {
                error = "复制幻灯片失败";
                return false;
            }

            created = dup[1];
            if (created.SlideIndex != toIndex)
            {
                created.MoveTo(toIndex);
            }

            return true;
        }

        private static bool TryCopyFromOne(
            PowerPoint.Slide source,
            PowerPoint.Presentation dest,
            int toIndex,
            out PowerPoint.Slide created,
            out string error)
        {
            created = null;
            error = null;
            PowerPoint.Application app = dest.Application;
            PowerPoint.PpAlertLevel prevAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
            try
            {
                try
                {
                    prevAlerts = app.DisplayAlerts;
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch (Exception)
                {
                }

                source.Copy();
                PowerPoint.SlideRange pasted = dest.Slides.Paste(toIndex);
                if (pasted == null || pasted.Count < 1)
                {
                    error = "跨文档粘贴幻灯片失败";
                    return false;
                }

                created = pasted[1];
                ApplyKeepSourceFormatting(source, created);
                if (created.SlideIndex != toIndex)
                {
                    created.MoveTo(toIndex);
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
                try
                {
                    app.DisplayAlerts = prevAlerts;
                }
                catch (Exception)
                {
                }
            }
        }

        private static void ApplyKeepSourceFormatting(PowerPoint.Slide source, PowerPoint.Slide dest)
        {
            try
            {
                dest.Design = source.Design;
            }
            catch (Exception)
            {
            }

            try
            {
                dest.FollowMasterBackground = source.FollowMasterBackground;
            }
            catch (Exception)
            {
            }

            try
            {
                dest.ColorScheme = source.ColorScheme;
            }
            catch (Exception)
            {
            }
        }

        private static bool AreSameApplication(
            PowerPoint.Presentation dest,
            PowerPoint.Presentation source,
            out string error)
        {
            error = null;
            try
            {
                int destHwnd = dest.Application.HWND;
                int sourceHwnd = source.Application.HWND;
                if (destHwnd != sourceHwnd)
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
            PowerPoint.Presentation presentation,
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

            var slides = new List<PowerPoint.Slide>(sequence.Length);
            foreach (string id in sequence)
            {
                if (!TryFindSlideById(presentation, id, out PowerPoint.Slide slide, out error))
                {
                    return false;
                }

                slides.Add(slide);
            }

            if (sequence.Length >= presentation.Slides.Count)
            {
                error = "不能删除演示文稿中唯一的幻灯片";
                return false;
            }

            PowerPoint.Application app = presentation.Application;
            PowerPoint.PpAlertLevel prevAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
            try
            {
                try
                {
                    prevAlerts = app.DisplayAlerts;
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch (Exception)
                {
                }

                for (int i = 0; i < slides.Count; i++)
                {
                    string id = sequence[i];
                    slides[i].Delete();
                    deleted.Add(id);
                    focusId = id;
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
                try
                {
                    app.DisplayAlerts = prevAlerts;
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool TryMove(
            PowerPoint.Presentation presentation,
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

            int count = presentation.Slides.Count;
            int toIndex = request.ToIndex.Value;
            if (toIndex < 1 || toIndex > count)
            {
                error = "to_index 越界（move 允许 1.." + count + "）";
                return false;
            }

            if (!TryFindSlideById(presentation, request.SlideId.Trim(), out PowerPoint.Slide slide, out error))
            {
                return false;
            }

            focusId = Convert.ToString(slide.SlideID) ?? request.SlideId.Trim();
            if (slide.SlideIndex == toIndex)
            {
                focusIndex = toIndex;
                return true;
            }

            slide.MoveTo(toIndex);
            focusIndex = slide.SlideIndex;
            return true;
        }

        private static bool TryReorder(
            PowerPoint.Presentation presentation,
            PresentationManageSlideRequest request,
            out string error)
        {
            error = null;
            if (request.Order == null || request.Order.Length == 0)
            {
                error = "reorder 须提供 order（全部 slide_id 的排列）";
                return false;
            }

            int count = presentation.Slides.Count;
            if (request.Order.Length != count)
            {
                error = "order 长度须等于当前页数 " + count;
                return false;
            }

            var currentIds = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 1; i <= count; i++)
            {
                currentIds.Add(Convert.ToString(presentation.Slides[i].SlideID) ?? "");
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
                if (!TryFindSlideById(presentation, wantId, out PowerPoint.Slide slide, out error))
                {
                    return false;
                }

                if (slide.SlideIndex != targetPos)
                {
                    slide.MoveTo(targetPos);
                }
            }

            return true;
        }

        private static bool TryResolveCustomLayout(
            PowerPoint.Presentation presentation,
            string layoutName,
            out PowerPoint.CustomLayout layout,
            out string error)
        {
            layout = null;
            error = null;

            PowerPoint.CustomLayouts layouts;
            try
            {
                layouts = presentation.SlideMaster.CustomLayouts;
            }
            catch (Exception ex)
            {
                error = "读取版式失败: " + ex.Message;
                return false;
            }

            if (layouts == null || layouts.Count < 1)
            {
                error = "演示文稿无可用版式";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(layoutName))
            {
                string want = layoutName.Trim();
                for (int i = 1; i <= layouts.Count; i++)
                {
                    PowerPoint.CustomLayout candidate = layouts[i];
                    string name = "";
                    try
                    {
                        name = candidate.Name ?? "";
                    }
                    catch (Exception)
                    {
                    }

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

            // 未指定：优先空白，否则第一个
            for (int i = 1; i <= layouts.Count; i++)
            {
                PowerPoint.CustomLayout candidate = layouts[i];
                string name = "";
                try
                {
                    name = candidate.Name ?? "";
                }
                catch (Exception)
                {
                }

                if (name.IndexOf("空白", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Blank", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    layout = candidate;
                    return true;
                }
            }

            layout = layouts[1];
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

        private static PresentationManageSlideResult BuildResult(
            string channelId,
            string sourceChannelId,
            string kind,
            string action,
            string focusId,
            int? focusIndex,
            PowerPoint.Presentation presentation,
            List<ClearedPlaceholderInfo> cleared,
            List<CreatedSlideInfo> created,
            List<string> deleted)
        {
            var slides = new List<PresentationSlideInfo>();
            int count = 0;
            try
            {
                count = presentation.Slides.Count;
                int limit = Math.Min(count, PresentationContentResult.MaxSlides);
                for (int i = 1; i <= limit; i++)
                {
                    PowerPoint.Slide s = presentation.Slides[i];
                    slides.Add(ReadLight(s, i));
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
                Slides = slides,
                ClearedPlaceholders = cleared,
                Created = created,
                Deleted = deleted
            };
        }

        private static PresentationSlideInfo ReadLight(PowerPoint.Slide slide, int fallbackIndex)
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
            }

            return info;
        }
    }
}
