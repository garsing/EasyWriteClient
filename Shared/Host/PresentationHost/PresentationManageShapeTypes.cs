using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationManageShapeRequest
    {
        public string Action { get; set; }

        public string SlideId { get; set; }

        public string ToSlideId { get; set; }

        public string[] ShapeIds { get; set; }

        public string ShapeId { get; set; }

        public double? LeftPct { get; set; }

        public double? TopPct { get; set; }

        public string SourceChannelId { get; set; }
    }

    public sealed class PresentationManageShapeResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string SlideId { get; set; }

        public int DeletedCount { get; set; }

        public List<string> DeletedShapeIds { get; set; }

        public string GroupShapeId { get; set; }

        public string SourceSlideId { get; set; }

        public string SourceChannelId { get; set; }

        public string SourceShapeId { get; set; }

        public List<string> Warnings { get; set; }
    }

    public sealed class ClearedPlaceholderInfo
    {
        public string PlaceholderType { get; set; }

        public int ShapeComId { get; set; }
    }
}
