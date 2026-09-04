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

        /// <summary>B4：left/center/right/justify</summary>
        public string Align { get; set; }

        /// <summary>I6：top/middle/bottom；null=未读到</summary>
        public string Valign { get; set; }

        /// <summary>I7：可见文字实测宽，相对幻灯片宽度百分比</summary>
        public double? TextWidthPct { get; set; }

        /// <summary>B4：行距倍数或 exact:N（pt）</summary>
        public string LineSpacing { get; set; }

        /// <summary>B4：段前 pt</summary>
        public double? SpaceBeforePt { get; set; }

        /// <summary>B4：段后 pt</summary>
        public double? SpaceAfterPt { get; set; }

        /// <summary>B4：左缩进 pt</summary>
        public double? IndentLeftPt { get; set; }

        /// <summary>B4：首行缩进 pt（负=悬挂）</summary>
        public double? IndentFirstPt { get; set; }

        /// <summary>B4：none / bullet / number</summary>
        public string Bullet { get; set; }

        /// <summary>文本框内边距 pt（缺省不输出）</summary>
        public double? MarginLeftPt { get; set; }

        public double? MarginRightPt { get; set; }

        public double? MarginTopPt { get; set; }

        public double? MarginBottomPt { get; set; }

        public string Name { get; set; }

        public double? Rotation { get; set; }

        public bool TextTruncated { get; set; }

        /// <summary>骨架到层空壳：本组还有更深子节点未写出</summary>
        public bool DepthCapped { get; set; }

        internal PptHtmlChartFormat ChartFormat { get; set; }

        public List<PptHtmlShapeNode> Children { get; set; }
    }

    public sealed class PptHtmlReadResult
    {
        public const int MaxShapes = 200;

        public const int MaxTextChars = 4000;

        public const int SkeletonMaxDepth = 3;

        public const int SkeletonTextMaxChars = 40;

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

        /// <summary>默认页 / 组 shape_id 的骨架窗。叶子详细读为 false。</summary>
        public bool IsSkeleton { get; set; }

        /// <summary>本窗因深度截断的 group ShapeId，树上遇见顺序。</summary>
        public List<string> DepthCappedShapeIds { get; set; }

        /// <summary>字色诊断文件（会话目录相对名，如 read_font_color_slide_439.txt）</summary>
        public string DebugFilename { get; set; }

        public List<string> DebugTrace { get; set; }
    }
}
