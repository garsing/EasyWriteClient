using System;
using System.Collections.Generic;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PptTransitionPowerPointManager
    {
        public static bool TryTransition(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
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
                error = "须提供 action（list|get|set|clear）";
                return false;
            }

            string action = request.Action.Trim().ToLowerInvariant();
            try
            {
                switch (action)
                {
                    case "list":
                        return TryList(presentation, channelId, request, out result, out error);
                    case "get":
                        return TryGet(presentation, channelId, request, out result, out error);
                    case "set":
                        return TrySet(presentation, channelId, request, out result, out error);
                    case "clear":
                        return TryClear(presentation, channelId, request, out result, out error);
                    default:
                        error = "action 须为 list|get|set|clear";
                        return false;
                }
            }
            catch (Exception ex)
            {
                error = "换页过渡失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryList(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            if (request.ListHasForbiddenFilter || !string.IsNullOrWhiteSpace(request.SlideId))
            {
                error = "list 不接受 slide_id / slide_ids（全表扫描；单页请用 get）";
                return false;
            }

            var transitions = new List<PresentationTransitionInfo>();
            var warnings = new List<string>();
            for (int i = 1; i <= presentation.Slides.Count; i++)
            {
                PowerPoint.Slide slide = presentation.Slides[i];
                transitions.Add(ReadOne(slide, i, warnings));
            }

            result = BuildResult(channelId, "ppt", "list", null, transitions, warnings);
            error = null;
            return true;
        }

        private static bool TryGet(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out PowerPoint.Slide slide, out string slideId, out error))
            {
                return false;
            }

            var warnings = new List<string>();
            var info = ReadOne(slide, SafeIndex(slide), warnings);
            result = BuildResult(
                channelId,
                "ppt",
                "get",
                slideId,
                new List<PresentationTransitionInfo> { info },
                warnings);
            return true;
        }

        private static bool TrySet(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out PowerPoint.Slide slide, out string slideId, out error))
            {
                return false;
            }

            if (!PptTransitionEffectMap.TryResolveToEntryEffect(
                    request.Effect,
                    request.Dir,
                    out int entryEffect,
                    out error))
            {
                return false;
            }

            if (request.DurationMs.HasValue && request.DurationMs.Value < 0)
            {
                error = "duration_ms 不能为负";
                return false;
            }

            if (request.AdvanceAfterMs.HasValue && request.AdvanceAfterMs.Value < 0)
            {
                error = "advance_after_ms 不能为负";
                return false;
            }

            PowerPoint.SlideShowTransition t = slide.SlideShowTransition;
            t.EntryEffect = PptTransitionEffectMap.ToPowerPoint(entryEffect);

            if (request.DurationMs.HasValue)
            {
                t.Duration = request.DurationMs.Value / 1000.0f;
            }

            if (request.AdvanceOnClick.HasValue)
            {
                t.AdvanceOnClick = request.AdvanceOnClick.Value
                    ? Office.MsoTriState.msoTrue
                    : Office.MsoTriState.msoFalse;
            }

            if (request.AdvanceAfterMs.HasValue)
            {
                t.AdvanceOnTime = Office.MsoTriState.msoTrue;
                t.AdvanceTime = request.AdvanceAfterMs.Value / 1000.0f;
            }

            var warnings = new List<string>();
            var info = ReadOne(slide, SafeIndex(slide), warnings);
            result = BuildResult(
                channelId,
                "ppt",
                "set",
                slideId,
                new List<PresentationTransitionInfo> { info },
                warnings);
            return true;
        }

        private static bool TryClear(
            PowerPoint.Presentation presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out PowerPoint.Slide slide, out string slideId, out error))
            {
                return false;
            }

            // 只清效果；不改 Advance* / Hidden
            slide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;

            var warnings = new List<string>();
            var info = ReadOne(slide, SafeIndex(slide), warnings);
            result = BuildResult(
                channelId,
                "ppt",
                "clear",
                slideId,
                new List<PresentationTransitionInfo> { info },
                warnings);
            return true;
        }

        private static PresentationTransitionInfo ReadOne(
            PowerPoint.Slide slide,
            int fallbackIndex,
            List<string> warnings)
        {
            var info = new PresentationTransitionInfo
            {
                SlideId = "",
                Index = fallbackIndex,
                Effect = "none",
                Dir = null,
                DurationMs = null,
                AdvanceOnClick = true,
                AdvanceAfterMs = null
            };

            try
            {
                info.SlideId = Convert.ToString(slide.SlideID) ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                info.Index = slide.SlideIndex;
            }
            catch (Exception)
            {
            }

            try
            {
                int entry = (int)slide.SlideShowTransition.EntryEffect;
                PptTransitionEffectMap.TryDescribe(entry, out string effect, out string dir, out string warn);
                info.Effect = effect ?? "unknown";
                info.Dir = dir;
                if (!string.IsNullOrEmpty(warn) && warnings != null)
                {
                    warnings.Add("slide_id=" + info.SlideId + ": " + warn);
                }
            }
            catch (Exception ex)
            {
                warnings?.Add("slide_id=" + info.SlideId + ": 读取 EntryEffect 失败: " + ex.Message);
            }

            try
            {
                float sec = slide.SlideShowTransition.Duration;
                info.DurationMs = (int)Math.Round(sec * 1000.0, MidpointRounding.AwayFromZero);
            }
            catch (Exception)
            {
            }

            try
            {
                info.AdvanceOnClick = slide.SlideShowTransition.AdvanceOnClick == Office.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
            }

            try
            {
                if (slide.SlideShowTransition.AdvanceOnTime == Office.MsoTriState.msoTrue)
                {
                    float sec = slide.SlideShowTransition.AdvanceTime;
                    info.AdvanceAfterMs = (int)Math.Round(sec * 1000.0, MidpointRounding.AwayFromZero);
                }
            }
            catch (Exception)
            {
            }

            return info;
        }

        private static bool TryGetSlide(
            PowerPoint.Presentation presentation,
            string slideIdRaw,
            out PowerPoint.Slide slide,
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

        private static int SafeIndex(PowerPoint.Slide slide)
        {
            try
            {
                return slide.SlideIndex;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static PresentationTransitionResult BuildResult(
            string channelId,
            string kind,
            string action,
            string slideId,
            List<PresentationTransitionInfo> transitions,
            List<string> warnings)
        {
            return new PresentationTransitionResult
            {
                ChannelId = channelId ?? "",
                Kind = kind,
                Action = action,
                SlideId = slideId ?? "",
                Transitions = transitions ?? new List<PresentationTransitionInfo>(),
                Warnings = warnings != null && warnings.Count > 0 ? warnings : null
            };
        }
    }
}
