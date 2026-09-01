using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationSlideInfo
    {
        public int Index { get; set; }

        /// <summary>COM Slide.SlideID；字符串避免 JSON 大数问题。</summary>
        public string SlideId { get; set; }

        public string Title { get; set; }

        public string Layout { get; set; }

        public bool Hidden { get; set; }

        /// <summary>null = 未探测；true/false = 已判断。</summary>
        public bool? HasNotes { get; set; }
    }

    public sealed class PresentationContentResult
    {
        public const int MaxSlides = 500;

        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Name { get; set; }

        public string Path { get; set; }

        /// <summary>COM Presentations/Slides.Count（截断前）。</summary>
        public int SlideCount { get; set; }

        public List<PresentationSlideInfo> Slides { get; set; }

        public bool Truncated { get; set; }

        public string TruncatedReason { get; set; }

        public List<string> Palette { get; set; }

        public bool PaletteSampled { get; set; }

        public int PaletteScannedCount { get; set; }
    }

    internal sealed class PresentationCaptureResult
    {
        public string ChannelId { get; set; }
        public string Kind { get; set; }
        public int PageNumber { get; set; }
        public int SlideCount { get; set; }
        public string SlideId { get; set; }
        public ImageCaptureCompressor.CompressedImage Image { get; set; }
    }
}
