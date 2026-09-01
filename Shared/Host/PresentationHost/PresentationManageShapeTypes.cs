using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationManageShapeRequest
    {
        public string Action { get; set; }

        public string SlideId { get; set; }

        public string[] ShapeIds { get; set; }
    }

    public sealed class PresentationManageShapeResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string SlideId { get; set; }

        public int DeletedCount { get; set; }

        public List<string> DeletedShapeIds { get; set; }

        public List<string> Warnings { get; set; }
    }

    public sealed class ClearedPlaceholderInfo
    {
        public string PlaceholderType { get; set; }

        public int ShapeComId { get; set; }
    }
}
