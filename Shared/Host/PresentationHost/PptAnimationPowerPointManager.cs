using System;
using System.Collections.Generic;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PptAnimationPowerPointManager
    {
        public static bool TryAnimate(
            PowerPoint.Presentation presentation,
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
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, required: true, out PowerPoint.Slide slide, out string slideId, out error))
            {
                return false;
            }

            int? filterComId = null;
            if (!string.IsNullOrWhiteSpace(request.ShapeId))
            {
                if (!TryValidateShapeId(request.ShapeId, slideId, out filterComId, out error))
                {
                    return false;
                }
            }

            var effects = ReadEffects(slide, slideId, filterComId, out List<string> warnings);
            result = BuildResult(channelId, "ppt", "list", slideId, effects, 0, 0, warnings);
            return true;
        }

        private static bool TryAdd(
            PowerPoint.Presentation presentation,
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

            string slideIdHint = request.SlideId;
            if (!PptShapeId.TryParseShape(spec.ShapeId.Trim(), out string sidFromShape, out int comId))
            {
                error = "非法 ShapeId: " + spec.ShapeId;
                return false;
            }

            if (!string.IsNullOrWhiteSpace(slideIdHint)
                && !string.Equals(slideIdHint.Trim(), sidFromShape, StringComparison.Ordinal))
            {
                error = "ShapeId 与 slide_id 不一致";
                return false;
            }

            if (!TryGetSlide(presentation, sidFromShape, required: true, out PowerPoint.Slide slide, out string slideId, out error))
            {
                return false;
            }

            if (!TryPlanOne(spec, slideId, out PlannedEffect planned, out error))
            {
                return false;
            }

            PowerPoint.Shape shape = FindShapeById(slide.Shapes, planned.ComId);
            if (shape == null)
            {
                error = "ShapeId 不存在: " + planned.ShapeId;
                return false;
            }

            if (!TryAddEffect(slide, shape, planned, out error))
            {
                return false;
            }

            var effects = ReadEffects(slide, slideId, null, out List<string> warnings);
            result = BuildResult(channelId, "ppt", "add", slideId, effects, 1, 0, warnings);
            return true;
        }

        private static bool TryReplaceAll(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, required: true, out PowerPoint.Slide slide, out string slideId, out error))
            {
                return false;
            }

            if (request.Effects == null)
            {
                error = "replace_all 须提供 effects（可为 []）";
                return false;
            }

            var planned = new List<PlannedEffect>();
            foreach (PresentationAnimationEffectSpec spec in request.Effects)
            {
                if (spec == null)
                {
                    continue;
                }

                if (!TryPlanOne(spec, slideId, out PlannedEffect one, out error))
                {
                    return false;
                }

                planned.Add(one);
            }

            foreach (PlannedEffect one in planned)
            {
                if (FindShapeById(slide.Shapes, one.ComId) == null)
                {
                    error = "ShapeId 不存在: " + one.ShapeId;
                    return false;
                }
            }

            int cleared = ClearSequence(slide.TimeLine.MainSequence, null);
            int added = 0;
            foreach (PlannedEffect one in planned)
            {
                PowerPoint.Shape shape = FindShapeById(slide.Shapes, one.ComId);
                if (!TryAddEffect(slide, shape, one, out error))
                {
                    return false;
                }

                added++;
            }

            var effects = ReadEffects(slide, slideId, null, out List<string> warnings);
            result = BuildResult(channelId, "ppt", "replace_all", slideId, effects, added, cleared, warnings);
            return true;
        }

        private static bool TryClear(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, required: true, out PowerPoint.Slide slide, out string slideId, out error))
            {
                return false;
            }

            int? filterComId = null;
            if (!string.IsNullOrWhiteSpace(request.ShapeId))
            {
                if (!TryValidateShapeId(request.ShapeId, slideId, out filterComId, out error))
                {
                    return false;
                }
            }

            int cleared = ClearSequence(slide.TimeLine.MainSequence, filterComId);
            var effects = ReadEffects(slide, slideId, filterComId, out List<string> warnings);
            result = BuildResult(channelId, "ppt", "clear", slideId, effects, 0, cleared, warnings);
            return true;
        }

        private static bool TryPlanOne(
            PresentationAnimationEffectSpec spec,
            string expectedSlideId,
            out PlannedEffect planned,
            out string error)
        {
            planned = null;
            if (spec == null || string.IsNullOrWhiteSpace(spec.ShapeId))
            {
                error = "效果缺少 ShapeId";
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

            if (!PptAnimationEffectMap.TryResolve(
                    spec.Category,
                    spec.Effect,
                    spec.Dir,
                    out PowerPoint.MsoAnimEffect effectId,
                    out bool isExit,
                    out bool needsDir,
                    out PowerPoint.MsoAnimDirection direction,
                    out error))
            {
                return false;
            }

            if (!PptAnimationEffectMap.TryParseTrigger(spec.Trigger, out PowerPoint.MsoAnimTriggerType trigger, out error))
            {
                return false;
            }

            planned = new PlannedEffect
            {
                ShapeId = spec.ShapeId.Trim(),
                SlideId = sid,
                ComId = comId,
                Category = (spec.Category ?? "").Trim().ToLowerInvariant(),
                Effect = (spec.Effect ?? "").Trim().ToLowerInvariant(),
                Trigger = string.IsNullOrWhiteSpace(spec.Trigger) ? "on_click" : spec.Trigger.Trim().ToLowerInvariant(),
                DurationMs = spec.DurationMs,
                DelayMs = spec.DelayMs ?? 0,
                Dir = needsDir ? (spec.Dir ?? "").Trim().ToLowerInvariant() : null,
                EffectId = effectId,
                IsExit = isExit,
                NeedsDir = needsDir,
                Direction = direction,
                TriggerType = trigger
            };
            return true;
        }

        private static bool TryAddEffect(
            PowerPoint.Slide slide,
            PowerPoint.Shape shape,
            PlannedEffect planned,
            out string error)
        {
            error = null;
            try
            {
                PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
                PowerPoint.Effect effect = sequence.AddEffect(
                    shape,
                    planned.EffectId,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    planned.TriggerType);

                if (planned.IsExit)
                {
                    effect.Exit = Office.MsoTriState.msoTrue;
                }

                if (planned.NeedsDir)
                {
                    try
                    {
                        effect.EffectParameters.Direction = planned.Direction;
                    }
                    catch (Exception)
                    {
                    }
                }

                try
                {
                    if (planned.DurationMs.HasValue)
                    {
                        effect.Timing.Duration = planned.DurationMs.Value / 1000f;
                    }

                    effect.Timing.TriggerDelayTime = planned.DelayMs / 1000f;
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

        private static int ClearSequence(PowerPoint.Sequence sequence, int? onlyComId)
        {
            if (sequence == null)
            {
                return 0;
            }

            int cleared = 0;
            for (int i = sequence.Count; i >= 1; i--)
            {
                PowerPoint.Effect effect = sequence[i];
                if (onlyComId.HasValue)
                {
                    try
                    {
                        if (effect.Shape == null || effect.Shape.Id != onlyComId.Value)
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
                    effect.Delete();
                    cleared++;
                }
                catch (Exception)
                {
                }
            }

            return cleared;
        }

        private static List<PresentationAnimationEffectSpec> ReadEffects(
            PowerPoint.Slide slide,
            string slideId,
            int? filterComId,
            out List<string> warnings)
        {
            warnings = new List<string>();
            var list = new List<PresentationAnimationEffectSpec>();
            PowerPoint.Sequence sequence;
            try
            {
                sequence = slide.TimeLine.MainSequence;
            }
            catch (Exception ex)
            {
                warnings.Add("读取 MainSequence 失败: " + ex.Message);
                return list;
            }

            for (int i = 1; i <= sequence.Count; i++)
            {
                try
                {
                    PowerPoint.Effect effect = sequence[i];
                    int comId = 0;
                    try
                    {
                        comId = effect.Shape.Id;
                    }
                    catch (Exception)
                    {
                        warnings.Add("跳过无 Shape 的效果 index=" + i);
                        continue;
                    }

                    if (filterComId.HasValue && comId != filterComId.Value)
                    {
                        continue;
                    }

                    PptAnimationEffectMap.DescribeEffect(effect, out string category, out string effectName, out string dir);
                    string trigger = "on_click";
                    int? durationMs = null;
                    int? delayMs = null;
                    try
                    {
                        trigger = PptAnimationEffectMap.TriggerToName(effect.Timing.TriggerType);
                        durationMs = (int)Math.Round(effect.Timing.Duration * 1000);
                        delayMs = (int)Math.Round(effect.Timing.TriggerDelayTime * 1000);
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

        private static bool TryGetSlide(
            PowerPoint.Presentation presentation,
            string slideIdRaw,
            bool required,
            out PowerPoint.Slide slide,
            out string slideId,
            out string error)
        {
            slide = null;
            slideId = null;
            error = null;
            if (string.IsNullOrWhiteSpace(slideIdRaw))
            {
                if (required)
                {
                    error = "须提供 slide_id";
                    return false;
                }

                return false;
            }

            slideId = slideIdRaw.Trim();
            if (!int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idInt))
            {
                error = "非法 slide_id: " + slideId;
                return false;
            }

            for (int i = 1; i <= presentation.Slides.Count; i++)
            {
                PowerPoint.Slide candidate = presentation.Slides[i];
                if (candidate.SlideID == idInt)
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

        private static PresentationAnimationEffectSpec FirstSpec(PresentationAnimationRequest request)
        {
            if (request.Effects != null && request.Effects.Count == 1)
            {
                return request.Effects[0];
            }

            if (!string.IsNullOrWhiteSpace(request.ShapeId)
                && request.Effects != null
                && request.Effects.Count > 0)
            {
                return null;
            }

            // 扁平字段：工具层会塞进 Effects[0]
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
                for (int i = 1; i <= items.Count; i++)
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

        private sealed class PlannedEffect
        {
            public string ShapeId;
            public string SlideId;
            public int ComId;
            public string Category;
            public string Effect;
            public string Trigger;
            public int? DurationMs;
            public int DelayMs;
            public string Dir;
            public PowerPoint.MsoAnimEffect EffectId;
            public bool IsExit;
            public bool NeedsDir;
            public PowerPoint.MsoAnimDirection Direction;
            public PowerPoint.MsoAnimTriggerType TriggerType;
        }
    }
}
