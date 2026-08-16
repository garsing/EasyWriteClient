using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationAnimationEffectSpec
    {
        public string ShapeId { get; set; }

        public string Category { get; set; }

        public string Effect { get; set; }

        public string Trigger { get; set; }

        public int? DurationMs { get; set; }

        public int? DelayMs { get; set; }

        public string Dir { get; set; }

        public int? Order { get; set; }

        public int? EffectIndex { get; set; }
    }

    public sealed class PresentationAnimationRequest
    {
        public string Action { get; set; }

        public string SlideId { get; set; }

        public string ShapeId { get; set; }

        public List<PresentationAnimationEffectSpec> Effects { get; set; }
    }

    public sealed class PresentationAnimationResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string SlideId { get; set; }

        public List<PresentationAnimationEffectSpec> Effects { get; set; }

        public int AddedCount { get; set; }

        public int ClearedCount { get; set; }

        public List<string> Warnings { get; set; }
    }
}
