using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>WPS 演示换页过渡（晚绑定；属性订不上则明确失败）。</summary>
    internal static class PptTransitionWppManager
    {
        private const int MsoTrue = -1;
        private const int MsoFalse = 0;

        public static bool TryTransition(
            object presentation,
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
            object presentation,
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

            if (!TryGetSlides(presentation, out object slides, out error))
            {
                return false;
            }

            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            var transitions = new List<PresentationTransitionInfo>();
            var warnings = new List<string>();
            for (int i = 1; i <= count; i++)
            {
                object slide = WppCom.GetIndexed(slides, i);
                if (slide == null)
                {
                    continue;
                }

                if (!TryReadOne(slide, i, out PresentationTransitionInfo info, out string readError))
                {
                    warnings.Add("index=" + i + ": " + readError);
                    continue;
                }

                transitions.Add(info);
            }

            result = BuildResult(channelId, "wpp", "list", null, transitions, warnings);
            error = null;
            return true;
        }

        private static bool TryGet(
            object presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out object slide, out string slideId, out error))
            {
                return false;
            }

            if (!TryReadOne(slide, 0, out PresentationTransitionInfo info, out error))
            {
                return false;
            }

            result = BuildResult(
                channelId,
                "wpp",
                "get",
                slideId,
                new List<PresentationTransitionInfo> { info },
                null);
            return true;
        }

        private static bool TrySet(
            object presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out object slide, out string slideId, out error))
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

            if (!TryGetTransition(slide, out object t, out error))
            {
                return false;
            }

            try
            {
                SetProperty(t, "EntryEffect", entryEffect);
            }
            catch (Exception ex)
            {
                error = "WPS 演示不支持设置 EntryEffect: " + ex.Message;
                return false;
            }

            if (request.DurationMs.HasValue)
            {
                try
                {
                    SetProperty(t, "Duration", request.DurationMs.Value / 1000.0);
                }
                catch (Exception ex)
                {
                    error = "WPS 演示不支持设置 Duration: " + ex.Message;
                    return false;
                }
            }

            if (request.AdvanceOnClick.HasValue)
            {
                try
                {
                    SetProperty(
                        t,
                        "AdvanceOnClick",
                        request.AdvanceOnClick.Value ? MsoTrue : MsoFalse);
                }
                catch (Exception ex)
                {
                    error = "WPS 演示不支持设置 AdvanceOnClick: " + ex.Message;
                    return false;
                }
            }

            if (request.AdvanceAfterMs.HasValue)
            {
                try
                {
                    SetProperty(t, "AdvanceOnTime", MsoTrue);
                    SetProperty(t, "AdvanceTime", request.AdvanceAfterMs.Value / 1000.0);
                }
                catch (Exception ex)
                {
                    error = "WPS 演示不支持设置 AdvanceOnTime/AdvanceTime: " + ex.Message;
                    return false;
                }
            }

            if (!TryReadOne(slide, 0, out PresentationTransitionInfo info, out error))
            {
                return false;
            }

            result = BuildResult(
                channelId,
                "wpp",
                "set",
                slideId,
                new List<PresentationTransitionInfo> { info },
                null);
            return true;
        }

        private static bool TryClear(
            object presentation,
            string channelId,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out string error)
        {
            result = null;
            if (!TryGetSlide(presentation, request.SlideId, out object slide, out string slideId, out error))
            {
                return false;
            }

            if (!TryGetTransition(slide, out object t, out error))
            {
                return false;
            }

            try
            {
                SetProperty(t, "EntryEffect", PptTransitionEffectMap.EffectNone);
            }
            catch (Exception ex)
            {
                error = "WPS 演示不支持清除 EntryEffect: " + ex.Message;
                return false;
            }

            if (!TryReadOne(slide, 0, out PresentationTransitionInfo info, out error))
            {
                return false;
            }

            result = BuildResult(
                channelId,
                "wpp",
                "clear",
                slideId,
                new List<PresentationTransitionInfo> { info },
                null);
            return true;
        }

        private static bool TryReadOne(
            object slide,
            int fallbackIndex,
            out PresentationTransitionInfo info,
            out string error)
        {
            info = new PresentationTransitionInfo
            {
                SlideId = "",
                Index = fallbackIndex,
                Effect = "none",
                Dir = null,
                DurationMs = null,
                AdvanceOnClick = true,
                AdvanceAfterMs = null
            };
            error = null;

            try
            {
                info.SlideId = Convert.ToString(WppCom.GetProperty(slide, "SlideID")) ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                info.Index = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex"));
            }
            catch (Exception)
            {
            }

            if (!TryGetTransition(slide, out object t, out error))
            {
                return false;
            }

            try
            {
                int entry = Convert.ToInt32(WppCom.GetProperty(t, "EntryEffect"));
                PptTransitionEffectMap.TryDescribe(entry, out string effect, out string dir, out string warn);
                info.Effect = effect ?? "unknown";
                info.Dir = dir;
                if (!string.IsNullOrEmpty(warn))
                {
                    // 读侧可带 warning，不失败
                }
            }
            catch (Exception ex)
            {
                error = "读取 EntryEffect 失败: " + ex.Message;
                return false;
            }

            try
            {
                double sec = Convert.ToDouble(WppCom.GetProperty(t, "Duration"), CultureInfo.InvariantCulture);
                info.DurationMs = (int)Math.Round(sec * 1000.0, MidpointRounding.AwayFromZero);
            }
            catch (Exception)
            {
            }

            try
            {
                int click = Convert.ToInt32(WppCom.GetProperty(t, "AdvanceOnClick"));
                info.AdvanceOnClick = click == MsoTrue || click == 1;
            }
            catch (Exception)
            {
            }

            try
            {
                int onTime = Convert.ToInt32(WppCom.GetProperty(t, "AdvanceOnTime"));
                if (onTime == MsoTrue || onTime == 1)
                {
                    double sec = Convert.ToDouble(WppCom.GetProperty(t, "AdvanceTime"), CultureInfo.InvariantCulture);
                    info.AdvanceAfterMs = (int)Math.Round(sec * 1000.0, MidpointRounding.AwayFromZero);
                }
            }
            catch (Exception)
            {
            }

            return true;
        }

        private static bool TryGetTransition(object slide, out object transition, out string error)
        {
            transition = null;
            error = null;
            try
            {
                transition = WppCom.GetProperty(slide, "SlideShowTransition");
            }
            catch (Exception ex)
            {
                error = "WPS 演示不支持 SlideShowTransition: " + ex.Message;
                return false;
            }

            if (transition == null)
            {
                error = "WPS 演示 SlideShowTransition 为空";
                return false;
            }

            return true;
        }

        private static bool TryGetSlides(object presentation, out object slides, out string error)
        {
            slides = null;
            error = null;
            try
            {
                slides = WppCom.GetProperty(presentation, "Slides");
            }
            catch (Exception ex)
            {
                error = "无法读取 Slides: " + ex.Message;
                return false;
            }

            if (slides == null)
            {
                error = "Slides 为空";
                return false;
            }

            return true;
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

            if (!TryGetSlides(presentation, out object slides, out error))
            {
                return false;
            }

            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object candidate = WppCom.GetIndexed(slides, i);
                if (candidate == null)
                {
                    continue;
                }

                try
                {
                    int sid = Convert.ToInt32(WppCom.GetProperty(candidate, "SlideID"));
                    if (sid == idInt)
                    {
                        slide = candidate;
                        return true;
                    }
                }
                catch (Exception)
                {
                }
            }

            error = "幻灯片不存在: slide_id=" + slideId;
            return false;
        }

        private static void SetProperty(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.SetProperty
                    | System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.Public,
                null,
                target,
                new object[] { value });
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
