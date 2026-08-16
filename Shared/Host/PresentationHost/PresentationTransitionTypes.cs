using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationTransitionInfo
    {
        public string SlideId { get; set; }

        public int Index { get; set; }

        public string Effect { get; set; }

        public string Dir { get; set; }

        public int? DurationMs { get; set; }

        public bool AdvanceOnClick { get; set; }

        public int? AdvanceAfterMs { get; set; }
    }

    public sealed class PresentationTransitionRequest
    {
        public string Action { get; set; }

        public string SlideId { get; set; }

        public string Effect { get; set; }

        public string Dir { get; set; }

        public int? DurationMs { get; set; }

        public bool? AdvanceOnClick { get; set; }

        public int? AdvanceAfterMs { get; set; }

        /// <summary>list 时若调用方误传了 slide_id / slide_ids，工具层应先拒；此标志供 Host 二次校验。</summary>
        public bool ListHasForbiddenFilter { get; set; }
    }

    public sealed class PresentationTransitionResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string SlideId { get; set; }

        public List<PresentationTransitionInfo> Transitions { get; set; }

        public List<string> Warnings { get; set; }
    }
}
