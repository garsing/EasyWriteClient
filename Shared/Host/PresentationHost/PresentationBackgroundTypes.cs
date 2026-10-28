using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationBackgroundInfo
    {
        public string SlideId { get; set; }

        public int Index { get; set; }

        public bool FollowMaster { get; set; }

        /// <summary>master | solid | picture | other</summary>
        public string Kind { get; set; }

        public string Color { get; set; }

        public string Path { get; set; }

        /// <summary>stretch | tile</summary>
        public string Fit { get; set; }

        public double? StretchLeft { get; set; }

        public double? StretchRight { get; set; }

        public double? StretchTop { get; set; }

        public double? StretchBottom { get; set; }

        public double? TileOffsetX { get; set; }

        public double? TileOffsetY { get; set; }

        public double? TileScaleX { get; set; }

        public double? TileScaleY { get; set; }

        public string TileAlign { get; set; }

        public string TileMirror { get; set; }
    }

    public sealed class PresentationBackgroundRequest
    {
        public string Action { get; set; }

        public string SlideId { get; set; }

        public bool ListHasForbiddenFilter { get; set; }

        /// <summary>list / get / clear 若带了写字段。</summary>
        public bool ReadActionHasWriteFields { get; set; }

        public string Kind { get; set; }

        public string Color { get; set; }

        /// <summary>已解析的本机绝对路径。仅 picture。</summary>
        public string PictureAbsolutePath { get; set; }

        /// <summary>回包里的 path，保持调用方传入的写法。</summary>
        public string PathEcho { get; set; }

        public string Fit { get; set; }

        public bool SawStretch { get; set; }

        public bool SawTile { get; set; }

        public double? StretchLeft { get; set; }

        public double? StretchRight { get; set; }

        public double? StretchTop { get; set; }

        public double? StretchBottom { get; set; }

        public double? TileOffsetX { get; set; }

        public double? TileOffsetY { get; set; }

        public double? TileScaleX { get; set; }

        public double? TileScaleY { get; set; }

        public string TileAlign { get; set; }

        public string TileMirror { get; set; }

        /// <summary>get 图片时的落盘目录（会话下 ppt_backgrounds）。</summary>
        public string ExportDirectory { get; set; }
    }

    public sealed class PresentationBackgroundResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string SlideId { get; set; }

        public List<PresentationBackgroundInfo> Backgrounds { get; set; }

        public List<string> Warnings { get; set; }
    }
}
