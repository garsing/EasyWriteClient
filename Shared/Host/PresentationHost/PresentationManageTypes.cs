using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationManageSlideRequest
    {
        public string Action { get; set; }

        public string SlideId { get; set; }

        /// <summary>1-based；null = add 末尾 / duplicate 源页之后。</summary>
        public int? ToIndex { get; set; }

        public string Layout { get; set; }

        public string[] Order { get; set; }

        public bool Confirm { get; set; }
    }

    public sealed class PresentationManageSlideResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string FocusSlideId { get; set; }

        public int? FocusIndex { get; set; }

        public int SlideCount { get; set; }

        public List<PresentationSlideInfo> Slides { get; set; }

        public List<ClearedPlaceholderInfo> ClearedPlaceholders { get; set; }
    }
}
