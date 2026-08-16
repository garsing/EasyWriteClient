using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>WPS 演示页内动画（晚绑定；API 不全则明确失败）。</summary>
    internal static class PptAnimationWppManager
    {
        // 与 PowerPoint PIA 枚举数值对齐
        private const int EffectAppear = 1;
        private const int EffectFly = 2;
        private const int EffectFade = 10;
        private const int EffectWipe = 22;
        private const int EffectFloat = 30;
        private const int EffectAscend = 39;
        private const int EffectGrowShrink = 59;
        private const int EffectSpin = 61;

        private const int TriggerOnClick = 1;
        private const int TriggerWithPrevious = 2;
        private const int TriggerAfterPrevious = 3;

        private const int DirUp = 1;
        private const int DirRight = 2;
        private const int DirDown = 3;
        private const int DirLeft = 4;

        private const int AnimateLevelNone = 0;

        public static bool TryAnimate(
            object presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
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
                error = "须提供 action（list|add|replace_all|clear）";
                return false;
            }

            string action = request.Action.Trim().ToLowerInvariant();
            try
            {
                switch (action)
                {
                    case "list":
                        return TryList(presentation, channelId, request, out result, out error);
                    case "add":
                        return TryAdd(presentation, channelId, request, out result, out error);
                    case "replace_all":
                        return TryReplaceAll(presentation, channelId, request, out result, out error);
                    case "clear":
                        return TryClear(presentation, channelId, request, out result, out error);
                    default:
                        error = "action 须为 list|add|replace_all|clear";
                        return false;
                }
            }
            catch (Exception ex)
            {
                error = "页内动画失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryList(
            object presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out object slide, out string slideId, out error))
            {
                return false;
            }

            int? filterComId = null;
            if (!string.IsNullOrWhiteSpace(request.ShapeId)
                && !TryValidateShapeId(request.ShapeId, slideId, out filterComId, out error))
            {
                return false;
            }

            if (!TryGetMainSequence(slide, out object sequence, out error))
            {
                return false;
            }

            var effects = ReadEffects(sequence, slideId, filterComId, out List<string> warnings);
            result = BuildResult(channelId, "wpp", "list", slideId, effects, 0, 0, warnings);
            return true;
        }

        private static bool TryAdd(
            object presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            PresentationAnimationEffectSpec spec = FirstSpec(request);
            if (spec == null || string.IsNullOrWhiteSpace(spec.ShapeId))
            {
                error = "add 须提供 ShapeId 与效果字段";
                return false;
            }

            if (!PptShapeId.TryParseShape(spec.ShapeId.Trim(), out string sidFromShape, out int comId))
            {
                error = "非法 ShapeId: " + spec.ShapeId;
                return false;
            }

            if (!string.IsNullOrWhiteSpace(request.SlideId)
                && !string.Equals(request.SlideId.Trim(), sidFromShape, StringComparison.Ordinal))
            {
                error = "ShapeId 与 slide_id 不一致";
                return false;
            }

            if (!TryGetSlide(presentation, sidFromShape, out object slide, out string slideId, out error))
            {
                return false;
            }

            if (!TryPlan(spec, slideId, out Planned planned, out error))
            {
                return false;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            object shape = FindShapeById(shapes, planned.ComId);
            if (shape == null)
            {
                error = "ShapeId 不存在: " + planned.ShapeId;
                return false;
            }

            if (!TryGetMainSequence(slide, out object sequence, out error))
            {
                return false;
            }

            if (!TryAddEffect(sequence, shape, planned, out error))
            {
                return false;
            }

            var effects = ReadEffects(sequence, slideId, null, out List<string> warnings);
            result = BuildResult(channelId, "wpp", "add", slideId, effects, 1, 0, warnings);
            return true;
        }

        private static bool TryReplaceAll(
            object presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out object slide, out string slideId, out error))
            {
                return false;
            }

            if (request.Effects == null)
            {
                error = "replace_all 须提供 effects（可为 []）";
                return false;
            }

            var plannedList = new List<Planned>();
            foreach (PresentationAnimationEffectSpec spec in request.Effects)
            {
                if (spec == null)
                {
                    continue;
                }

                if (!TryPlan(spec, slideId, out Planned one, out error))
                {
                    return false;
                }

                plannedList.Add(one);
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            foreach (Planned one in plannedList)
            {
                if (FindShapeById(shapes, one.ComId) == null)
                {
                    error = "ShapeId 不存在: " + one.ShapeId;
                    return false;
                }
            }

            if (!TryGetMainSequence(slide, out object sequence, out error))
            {
                return false;
            }

            int cleared = ClearSequence(sequence, null);
            int added = 0;
            foreach (Planned one in plannedList)
            {
                object shape = FindShapeById(shapes, one.ComId);
                if (!TryAddEffect(sequence, shape, one, out error))
                {
                    return false;
                }

                added++;
            }

            var effects = ReadEffects(sequence, slideId, null, out List<string> warnings);
            result = BuildResult(channelId, "wpp", "replace_all", slideId, effects, added, cleared, warnings);
            return true;
        }

        private static bool TryClear(
            object presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out object slide, out string slideId, out error))
            {
                return false;
            }

            int? filterComId = null;
            if (!string.IsNullOrWhiteSpace(request.ShapeId)
                && !TryValidateShapeId(request.ShapeId, slideId, out filterComId, out error))
            {
                return false;
            }

            if (!TryGetMainSequence(slide, out object sequence, out error))
            {
                return false;
            }

            int cleared = ClearSequence(sequence, filterComId);
            var effects = ReadEffects(sequence, slideId, filterComId, out List<string> warnings);
            result = BuildResult(channelId, "wpp", "clear", slideId, effects, 0, cleared, warnings);
            return true;
        }

        private static bool TryPlan(
            PresentationAnimationEffectSpec spec,
            string expectedSlideId,
            out Planned planned,
            out string error)
        {
            planned = null;
            // 复用 PPT 映射的校验逻辑，再落到整型常量
            if (!PptAnimationEffectMap.TryResolve(
                    spec.Category,
                    spec.Effect,
                    spec.Dir,
                    out var effectEnum,
                    out bool isExit,
                    out bool needsDir,
                    out var dirEnum,
                    out error))
            {
                return false;
            }

            if (!PptAnimationEffectMap.TryParseTrigger(spec.Trigger, out var triggerEnum, out error))
            {
                return false;
            }

            if (!PptShapeId.TryParseShape(spec.ShapeId.Trim(), out string sid, out int comId))
            {
                error = "非法 ShapeId: " + spec.ShapeId;
                return false;
            }

            if (!string.Equals(sid, expectedSlideId, StringComparison.Ordinal))
            {
                error = "effects 中 ShapeId 须属于 slide_id=" + expectedSlideId;
                return false;
            }

            if (spec.DurationMs.HasValue && spec.DurationMs.Value < 0)
            {
                error = "duration_ms 不能为负";
                return false;
            }

            if (spec.DelayMs.HasValue && spec.DelayMs.Value < 0)
            {
                error = "delay_ms 不能为负";
                return false;
            }

            planned = new Planned
            {
                ShapeId = spec.ShapeId.Trim(),
                ComId = comId,
                EffectId = (int)effectEnum,
                IsExit = isExit,
                NeedsDir = needsDir,
                Direction = (int)dirEnum,
                TriggerType = (int)triggerEnum,
                DurationMs = spec.DurationMs,
                DelayMs = spec.DelayMs ?? 0
            };
            return true;
        }

        private static bool TryAddEffect(object sequence, object shape, Planned planned, out string error)
        {
            error = null;
            try
            {
                object effect = WppCom.Invoke(
                    sequence,
                    "AddEffect",
                    shape,
                    planned.EffectId,
                    AnimateLevelNone,
                    planned.TriggerType);
                if (effect == null)
                {
                    error = "添加动画失败";
                    return false;
                }

                if (planned.IsExit)
                {
                    WppCom.TrySetProperty(effect, "Exit", -1);
                }

                if (planned.NeedsDir)
                {
                    try
                    {
                        object parameters = WppCom.GetProperty(effect, "EffectParameters");
                        WppCom.TrySetProperty(parameters, "Direction", planned.Direction);
                    }
                    catch (Exception)
                    {
                    }
                }

                try
                {
                    object timing = WppCom.GetProperty(effect, "Timing");
                    if (planned.DurationMs.HasValue)
                    {
                        WppCom.TrySetProperty(timing, "Duration", planned.DurationMs.Value / 1000.0);
                    }

                    WppCom.TrySetProperty(timing, "TriggerDelayTime", planned.DelayMs / 1000.0);
                }
                catch (Exception)
                {
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "添加动画失败: " + ex.Message;
                return false;
            }
        }

        private static int ClearSequence(object sequence, int? onlyComId)
        {
            int cleared = 0;
            int count = Convert.ToInt32(WppCom.GetProperty(sequence, "Count"));
            for (int i = count; i >= 1; i--)
            {
                object effect = WppCom.GetIndexed(sequence, i);
                if (onlyComId.HasValue)
                {
                    try
                    {
                        object shape = WppCom.GetProperty(effect, "Shape");
                        int id = Convert.ToInt32(WppCom.GetProperty(shape, "Id"));
                        if (id != onlyComId.Value)
                        {
                            continue;
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                try
                {
                    WppCom.Invoke(effect, "Delete");
                    cleared++;
                }
                catch (Exception)
                {
                }
            }

            return cleared;
        }

        private static List<PresentationAnimationEffectSpec> ReadEffects(
            object sequence,
            string slideId,
            int? filterComId,
            out List<string> warnings)
        {
            warnings = new List<string>();
            var list = new List<PresentationAnimationEffectSpec>();
            int count = Convert.ToInt32(WppCom.GetProperty(sequence, "Count"));
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    object effect = WppCom.GetIndexed(sequence, i);
                    object shape = WppCom.GetProperty(effect, "Shape");
                    int comId = Convert.ToInt32(WppCom.GetProperty(shape, "Id"));
                    if (filterComId.HasValue && comId != filterComId.Value)
                    {
                        continue;
                    }

                    int effectType = Convert.ToInt32(WppCom.GetProperty(effect, "EffectType"));
                    bool isExit = false;
                    try
                    {
                        object exit = WppCom.GetProperty(effect, "Exit");
                        isExit = exit != null && (Convert.ToInt32(exit) == -1 || Convert.ToBoolean(exit));
                    }
                    catch (Exception)
                    {
                    }

                    MapEffect(effectType, isExit, out string category, out string effectName);
                    string trigger = "on_click";
                    int? durationMs = null;
                    int? delayMs = null;
                    string dir = null;
                    try
                    {
                        object timing = WppCom.GetProperty(effect, "Timing");
                        int trig = Convert.ToInt32(WppCom.GetProperty(timing, "TriggerType"));
                        trigger = trig == TriggerWithPrevious
                            ? "with_previous"
                            : trig == TriggerAfterPrevious
                                ? "after_previous"
                                : "on_click";
                        durationMs = (int)Math.Round(Convert.ToDouble(WppCom.GetProperty(timing, "Duration")) * 1000);
                        delayMs = (int)Math.Round(Convert.ToDouble(WppCom.GetProperty(timing, "TriggerDelayTime")) * 1000);
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        object parameters = WppCom.GetProperty(effect, "EffectParameters");
                        int d = Convert.ToInt32(WppCom.GetProperty(parameters, "Direction"));
                        dir = d == DirLeft ? "left"
                            : d == DirRight ? "right"
                            : d == DirUp ? "top"
                            : d == DirDown ? "bottom"
                            : null;
                    }
                    catch (Exception)
                    {
                    }

                    list.Add(new PresentationAnimationEffectSpec
                    {
                        ShapeId = PptShapeId.FormatShape(slideId, comId),
                        Category = category,
                        Effect = effectName,
                        Trigger = trigger,
                        DurationMs = durationMs,
                        DelayMs = delayMs,
                        Dir = dir,
                        Order = list.Count + 1,
                        EffectIndex = i
                    });
                }
                catch (Exception ex)
                {
                    warnings.Add("读取效果失败 index=" + i + ": " + ex.Message);
                }
            }

            return list;
        }

        private static void MapEffect(int effectType, bool isExit, out string category, out string effectName)
        {
            switch (effectType)
            {
                case EffectAppear:
                    effectName = "appear";
                    category = isExit ? "exit" : "entrance";
                    break;
                case EffectFade:
                    effectName = "fade";
                    category = isExit ? "exit" : "entrance";
                    break;
                case EffectFly:
                    effectName = isExit ? "fly_out" : "fly_in";
                    category = isExit ? "exit" : "entrance";
                    break;
                case EffectWipe:
                    effectName = "wipe";
                    category = isExit ? "exit" : "entrance";
                    break;
                case EffectAscend:
                case EffectFloat:
                    effectName = "float_up";
                    category = "entrance";
                    break;
                case EffectGrowShrink:
                    effectName = "grow_shrink";
                    category = "emphasis";
                    break;
                case EffectSpin:
                    effectName = "spin";
                    category = "emphasis";
                    break;
                default:
                    effectName = "appear";
                    category = isExit ? "exit" : "entrance";
                    break;
            }
        }

        private static bool TryGetMainSequence(object slide, out object sequence, out string error)
        {
            sequence = null;
            error = null;
            try
            {
                object timeline = WppCom.GetProperty(slide, "TimeLine");
                sequence = WppCom.GetProperty(timeline, "MainSequence");
                if (sequence == null)
                {
                    error = "无法读取 MainSequence（wpp 可能不支持页内动画 API）";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "无法读取 MainSequence: " + ex.Message;
                return false;
            }
        }

        private static bool TryGetSlide(
            object presentation,
            string slideIdRaw,
            out object slide,
            out string slideId,
            out string error)
        {
            slide = null;
            slideId = null;
            error = null;
            if (string.IsNullOrWhiteSpace(slideIdRaw))
            {
                error = "须提供 slide_id";
                return false;
            }

            slideId = slideIdRaw.Trim();
            if (!int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idInt))
            {
                error = "非法 slide_id: " + slideId;
                return false;
            }

            object slides = WppCom.GetProperty(presentation, "Slides");
            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object candidate = WppCom.GetIndexed(slides, i);
                int sid = Convert.ToInt32(WppCom.GetProperty(candidate, "SlideID"));
                if (sid == idInt)
                {
                    slide = candidate;
                    return true;
                }
            }

            error = "幻灯片不存在: slide_id=" + slideId;
            return false;
        }

        private static bool TryValidateShapeId(string shapeId, string slideId, out int? comId, out string error)
        {
            comId = null;
            if (!PptShapeId.TryParseShape(shapeId.Trim(), out string sid, out int id))
            {
                error = "非法 ShapeId: " + shapeId;
                return false;
            }

            if (!string.Equals(sid, slideId, StringComparison.Ordinal))
            {
                error = "ShapeId 与 slide_id 不一致";
                return false;
            }

            comId = id;
            error = null;
            return true;
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
                    int shapeId = Convert.ToInt32(WppCom.GetProperty(shape, "Id"));
                    if (shapeId == id)
                    {
                        return shape;
                    }

                    try
                    {
                        int type = Convert.ToInt32(WppCom.GetProperty(shape, "Type"));
                        if (type == 6)
                        {
                            object found = FindInGroup(shape, id);
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

        private static object FindInGroup(object group, int id)
        {
            try
            {
                object items = WppCom.GetProperty(group, "GroupItems");
                int count = Convert.ToInt32(WppCom.GetProperty(items, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object child = WppCom.GetIndexed(items, i);
                    if (Convert.ToInt32(WppCom.GetProperty(child, "Id")) == id)
                    {
                        return child;
                    }

                    try
                    {
                        if (Convert.ToInt32(WppCom.GetProperty(child, "Type")) == 6)
                        {
                            object nested = FindInGroup(child, id);
                            if (nested != null)
                            {
                                return nested;
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

        private static PresentationAnimationEffectSpec FirstSpec(PresentationAnimationRequest request)
        {
            if (request.Effects != null && request.Effects.Count > 0)
            {
                return request.Effects[0];
            }

            return null;
        }

        private static PresentationAnimationResult BuildResult(
            string channelId,
            string kind,
            string action,
            string slideId,
            List<PresentationAnimationEffectSpec> effects,
            int added,
            int cleared,
            List<string> warnings)
        {
            return new PresentationAnimationResult
            {
                ChannelId = channelId ?? "",
                Kind = kind,
                Action = action,
                SlideId = slideId ?? "",
                Effects = effects ?? new List<PresentationAnimationEffectSpec>(),
                AddedCount = added,
                ClearedCount = cleared,
                Warnings = warnings
            };
        }

        private sealed class Planned
        {
            public string ShapeId;
            public int ComId;
            public int EffectId;
            public bool IsExit;
            public bool NeedsDir;
            public int Direction;
            public int TriggerType;
            public int? DurationMs;
            public int DelayMs;
        }
    }
}
