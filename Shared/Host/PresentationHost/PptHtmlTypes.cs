using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PptHtmlShapeNode
    {
        public string ShapeId { get; set; }

        public string ShapeType { get; set; }

        /// <summary>h1 / div / img / table</summary>
        public string Tag { get; set; }

        public string Style { get; set; }

        public string Text { get; set; }

        /// <summary>table 时用；已是内部 HTML（无外层 table 标签时由序列化包）。</summary>
        public string InnerHtml { get; set; }

        public bool Editable { get; set; } = true;

        /// <summary>图片等文件源，如 workspace:ppt_images/a1b2c3d4.png（内容 hash 前 8 位）</summary>
        public string DataSrc { get; set; }

        /// <summary>B2：由 freeform/smartart/group/unknown 栅格而来时记原 type</summary>
        public string RasterizedFrom { get; set; }

        /// <summary>B1：#RRGGBB 或 none；null=未读到/不输出</summary>
        public string Fill { get; set; }

        /// <summary>B1：#RRGGBB；null=未读到/不输出</summary>
        public string FontColor { get; set; }

        /// <summary>B3：字号 pt；null=未读到</summary>
        public double? FontSizePt { get; set; }

        /// <summary>B3：加粗；null=未读到</summary>
        public bool? FontBold { get; set; }

        /// <summary>B3c：字体名（如 微软雅黑）；null=未读到</summary>
        public string FontName { get; set; }

        /// <summary>B3：叠放；数值越大越靠上（≈ ZOrderPosition）</summary>
        public int? Z { get; set; }

        /// <summary>B3b：#RRGGBB 或 none</summary>
        public string LineColor { get; set; }

        /// <summary>B3b：线宽 pt</summary>
        public double? LineWidthPt { get; set; }

        public string Name { get; set; }

        public double? Rotation { get; set; }

        public bool TextTruncated { get; set; }
    }

    public sealed class PptHtmlReadResult
    {
        public const int MaxShapes = 200;

        public const int MaxTextChars = 4000;

        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Name { get; set; }

        public string SlideId { get; set; }

        public int Index { get; set; }

        public string Layout { get; set; }

        public bool Hidden { get; set; }

        public bool? HasNotes { get; set; }

        public List<PptHtmlShapeNode> Shapes { get; set; }

        public bool Truncated { get; set; }

        public string TruncatedReason { get; set; }

        public int ShapeCount { get; set; }
    }
}
