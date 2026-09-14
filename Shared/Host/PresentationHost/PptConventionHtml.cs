using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace WordAddIn1.PresentationHost
{
    internal static class PptConventionHtml
    {
        public static string BuildDisplayContents(PptHtmlReadResult result)
        {
            if (result == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            sb.AppendLine("演示文稿：" + (result.Name ?? ""));
            sb.AppendLine(
                "kind=" + (result.Kind ?? "")
                + " channel_id=" + (result.ChannelId ?? ""));
            sb.Append("slide_id=").Append(result.SlideId ?? "");
            sb.Append(" index=").Append(result.Index);
            sb.Append(" layout=").Append(result.Layout ?? "");
            sb.Append(" hidden=").Append(result.Hidden ? "true" : "false");
            if (result.HasNotes.HasValue)
            {
                sb.Append(" has_notes=").Append(result.HasNotes.Value ? "true" : "false");
            }

            sb.AppendLine();
            if (result.Truncated)
            {
                sb.AppendLine(
                    "truncated=true"
                    + (string.IsNullOrEmpty(result.TruncatedReason)
                        ? ""
                        : " reason=" + result.TruncatedReason));
            }

            sb.AppendLine();
            if (result.IsSkeleton)
            {
                AppendSkeletonPreface(sb, result);
                sb.AppendLine();
            }

            sb.Append(BuildSectionHtml(result));
            return sb.ToString().TrimEnd();
        }

        public static List<string> CollectDepthCappedIds(IList<PptHtmlShapeNode> nodes)
        {
            var ids = new List<string>();
            CollectDepthCappedIds(nodes, ids);
            return ids;
        }

        public const double DefaultContentMinArea = 1.0;

        /// <summary>
        /// 精简骨架：去掉带 data-rasterized-from 的栅格装饰，并丢掉因此变空的组。
        /// </summary>
        public static void ApplyCompactFilter(List<PptHtmlShapeNode> shapes)
        {
            if (shapes == null)
            {
                return;
            }

            List<PptHtmlShapeNode> kept = FilterCompactNodes(shapes);
            shapes.Clear();
            shapes.AddRange(kept);
        }

        public static List<PptHtmlShapeNode> FilterCompactNodes(IList<PptHtmlShapeNode> nodes)
        {
            var kept = new List<PptHtmlShapeNode>();
            if (nodes == null)
            {
                return kept;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode node = FilterCompactNode(nodes[i]);
                if (node != null)
                {
                    kept.Add(node);
                }
            }

            return kept;
        }

        private static PptHtmlShapeNode FilterCompactNode(PptHtmlShapeNode node)
        {
            if (node == null)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(node.RasterizedFrom))
            {
                return null;
            }

            if (node.Children != null && node.Children.Count > 0)
            {
                node.Children = FilterCompactNodes(node.Children);
            }

            if (string.Equals(node.ShapeType, "group", StringComparison.Ordinal)
                && (node.Children == null || node.Children.Count == 0)
                && !node.DepthCapped
                && string.IsNullOrEmpty(node.Text))
            {
                return null;
            }

            return node;
        }

        /// <summary>
        /// 内容骨架：只留有字/textbox/占位符/chart/table/media，以及足够大的真图；空组丢掉。
        /// </summary>
        public static void ApplyContentOnlyFilter(List<PptHtmlShapeNode> shapes, double minArea)
        {
            if (shapes == null)
            {
                return;
            }

            List<PptHtmlShapeNode> kept = FilterContentOnlyNodes(shapes, minArea);
            shapes.Clear();
            shapes.AddRange(kept);
        }

        public static bool IsContentCarrier(PptHtmlShapeNode node, bool picturesNeedArea, double minArea)
        {
            if (node == null)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(node.RasterizedFrom))
            {
                return false;
            }

            string type = node.ShapeType ?? "";
            if (type == "group"
                || type == "chart"
                || type == "table"
                || type == "media"
                || type == "textbox"
                || type.StartsWith("placeholder_", StringComparison.Ordinal))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(node.Text))
            {
                return true;
            }

            if (type == "picture")
            {
                if (!picturesNeedArea)
                {
                    return true;
                }

                double area = PptHtmlNodeSearch.ComputeArea(node);
                if (area <= 0)
                {
                    return true;
                }

                return area + 0.0000001 >= minArea;
            }

            return false;
        }

        private static List<PptHtmlShapeNode> FilterContentOnlyNodes(IList<PptHtmlShapeNode> nodes, double minArea)
        {
            var kept = new List<PptHtmlShapeNode>();
            if (nodes == null)
            {
                return kept;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode node = FilterContentOnlyNode(nodes[i], minArea);
                if (node != null)
                {
                    kept.Add(node);
                }
            }

            return kept;
        }

        private static PptHtmlShapeNode FilterContentOnlyNode(PptHtmlShapeNode node, double minArea)
        {
            if (node == null)
            {
                return null;
            }

            if (node.Children != null && node.Children.Count > 0)
            {
                node.Children = FilterContentOnlyNodes(node.Children, minArea);
            }

            if (string.Equals(node.ShapeType, "group", StringComparison.Ordinal))
            {
                if ((node.Children == null || node.Children.Count == 0)
                    && !node.DepthCapped
                    && string.IsNullOrWhiteSpace(node.Text))
                {
                    return null;
                }

                return node;
            }

            if (!IsContentCarrier(node, picturesNeedArea: true, minArea))
            {
                return null;
            }

            return node;
        }

        public static string TruncateText(string text, int maxChars, out bool truncated)
        {
            truncated = false;
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            int limit = maxChars < 0 ? 0 : maxChars;
            if (normalized.Length <= limit)
            {
                return normalized;
            }

            truncated = true;
            return normalized.Substring(0, limit);
        }

        public static string BuildSectionHtml(PptHtmlReadResult result)
        {
            if (result == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            // section 不绑某一页 SlideID；apply 由工具参数 slide_id 指定目标页
            sb.Append("<section");
            sb.Append(" data-slide-index=\"").Append(result.Index).Append("\"");
            if (!string.IsNullOrEmpty(result.Layout))
            {
                sb.Append(" data-layout=\"").Append(EscapeAttr(result.Layout)).Append("\"");
            }

            sb.Append(" data-hidden=\"").Append(result.Hidden ? "true" : "false").Append("\"");
            if (result.HasNotes.HasValue)
            {
                sb.Append(" data-has-notes=\"")
                    .Append(result.HasNotes.Value ? "true" : "false")
                    .Append("\"");
            }

            if (result.Truncated)
            {
                sb.Append(" data-truncated=\"true\"");
            }

            sb.Append(" style=\"position:relative;width:100%;height:100%\"");
            sb.AppendLine(">");

            if (result.Shapes != null)
            {
                foreach (PptHtmlShapeNode node in result.Shapes)
                {
                    if (node == null)
                    {
                        continue;
                    }

                    AppendShape(sb, node);
                }
            }

            sb.Append("</section>");
            return sb.ToString();
        }

        public static string BuildFragment(IList<PptHtmlShapeNode> nodes)
        {
            var sb = new StringBuilder();
            if (nodes != null)
            {
                foreach (PptHtmlShapeNode node in nodes)
                {
                    if (node != null)
                    {
                        AppendShape(sb, node);
                    }
                }
            }

            return sb.ToString().TrimEnd();
        }

        public static string BuildStyle(double leftPct, double topPct, double widthPct, double heightPct)
        {
            return "position:absolute;left:" + FormatPct(leftPct)
                + "%;top:" + FormatPct(topPct)
                + "%;width:" + FormatPct(widthPct)
                + "%;height:" + FormatPct(heightPct) + "%";
        }

        public static string TruncateText(string text, out bool truncated)
        {
            return TruncateText(text, PptHtmlReadResult.MaxTextChars, out truncated);
        }

        public static string TruncateSkeletonText(string text, out bool truncated)
        {
            return TruncateText(text, PptHtmlReadResult.SkeletonTextMaxChars, out truncated);
        }

        private static void AppendShape(StringBuilder sb, PptHtmlShapeNode node)
        {
            AppendShape(sb, node, "  ");
        }

        private static void AppendShape(StringBuilder sb, PptHtmlShapeNode node, string indent)
        {
            string tag = string.IsNullOrEmpty(node.Tag) ? "div" : node.Tag;
            bool hasKids = node.Children != null && node.Children.Count > 0;
            if (tag == "img" && !hasKids)
            {
                sb.Append(indent).Append("<img");
                AppendShapeId(sb, node);
                AppendCommonAttrs(sb, node);
                sb.AppendLine(" />");
                return;
            }

            if (tag == "table" && !hasKids)
            {
                sb.Append(indent).Append("<table");
                AppendShapeId(sb, node);
                AppendCommonAttrs(sb, node);
                sb.AppendLine(">");
                if (!string.IsNullOrEmpty(node.InnerHtml))
                {
                    sb.Append(node.InnerHtml);
                    if (!node.InnerHtml.EndsWith("\n", StringComparison.Ordinal))
                    {
                        sb.AppendLine();
                    }
                }
                else if (!string.IsNullOrEmpty(node.Text))
                {
                    sb.Append(indent).Append("  ").Append(EscapeText(node.Text)).AppendLine();
                }

                sb.Append(indent).AppendLine("</table>");
                return;
            }

            if (node.ShapeType == "chart"
                && !hasKids
                && (node.ChartFormat != null || !string.IsNullOrEmpty(node.InnerHtml)))
            {
                sb.Append(indent).Append("<div");
                AppendShapeId(sb, node);
                AppendCommonAttrs(sb, node);
                if (node.ChartFormat != null)
                {
                    PptHtmlChartIo.AppendFormatAttrs(sb, node.ChartFormat);
                }

                sb.AppendLine(">");
                sb.Append(indent).AppendLine("  <table>");
                if (!string.IsNullOrEmpty(node.InnerHtml))
                {
                    sb.Append(node.InnerHtml);
                    if (!node.InnerHtml.EndsWith("\n", StringComparison.Ordinal))
                    {
                        sb.AppendLine();
                    }
                }

                sb.Append(indent).AppendLine("  </table>");
                sb.Append(indent).AppendLine("</div>");
                return;
            }

            sb.Append(indent).Append("<").Append(tag);
            AppendShapeId(sb, node);
            AppendCommonAttrs(sb, node);
            sb.Append(">");
            if (hasKids || node.ShapeType == "group")
            {
                sb.AppendLine();
                if (!string.IsNullOrEmpty(node.Text))
                {
                    sb.Append(indent).Append("  ").Append(EscapeText(node.Text)).AppendLine();
                }

                if (hasKids)
                {
                    foreach (PptHtmlShapeNode child in node.Children)
                    {
                        if (child != null)
                        {
                            AppendShape(sb, child, indent + "  ");
                        }
                    }
                }

                sb.Append(indent).Append("</").Append(tag).AppendLine(">");
                return;
            }

            sb.Append(EscapeText(node.Text));
            sb.Append("</").Append(tag).AppendLine(">");
        }

        private static void AppendShapeId(StringBuilder sb, PptHtmlShapeNode node)
        {
            if (!string.IsNullOrEmpty(node.ShapeId))
            {
                sb.Append(" ShapeId=\"").Append(EscapeAttr(node.ShapeId)).Append("\"");
            }
        }

        private static void AppendCommonAttrs(StringBuilder sb, PptHtmlShapeNode node)
        {
            if (!string.IsNullOrEmpty(node.ShapeType))
            {
                sb.Append(" data-shape-type=\"").Append(EscapeAttr(node.ShapeType)).Append("\"");
            }

            if (!node.Editable)
            {
                sb.Append(" data-editable=\"false\"");
            }

            if (!string.IsNullOrEmpty(node.DataSrc))
            {
                sb.Append(" data-src=\"").Append(EscapeAttr(node.DataSrc)).Append("\"");
            }

            if (!string.IsNullOrEmpty(node.RasterizedFrom))
            {
                sb.Append(" data-rasterized-from=\"")
                    .Append(EscapeAttr(node.RasterizedFrom))
                    .Append("\"");
            }

            if (node.SearchArea.HasValue)
            {
                sb.Append(" data-area=\"")
                    .Append(node.SearchArea.Value.ToString("0.0000", CultureInfo.InvariantCulture))
                    .Append("\"");
            }

            if (node.SearchAreaRank.HasValue)
            {
                sb.Append(" data-area-rank=\"")
                    .Append(node.SearchAreaRank.Value.ToString(CultureInfo.InvariantCulture))
                    .Append("\"");
            }

            if (!string.IsNullOrEmpty(node.Fill))
            {
                sb.Append(" data-fill=\"").Append(EscapeAttr(node.Fill)).Append("\"");
            }

            if (!string.IsNullOrEmpty(node.FontColor))
            {
                sb.Append(" data-font-color=\"").Append(EscapeAttr(node.FontColor)).Append("\"");
            }

            if (node.FontSizePt.HasValue)
            {
                sb.Append(" data-font-size=\"")
                    .Append(PptHtmlStyleIo.FormatPt(node.FontSizePt.Value))
                    .Append("\"");
            }

            if (node.FontBold.HasValue)
            {
                sb.Append(" data-font-bold=\"")
                    .Append(node.FontBold.Value ? "true" : "false")
                    .Append("\"");
            }

            if (!string.IsNullOrEmpty(node.FontName))
            {
                sb.Append(" data-font-name=\"").Append(EscapeAttr(node.FontName)).Append("\"");
            }

            if (node.Z.HasValue)
            {
                sb.Append(" data-z=\"")
                    .Append(node.Z.Value.ToString(CultureInfo.InvariantCulture))
                    .Append("\"");
            }

            if (!string.IsNullOrEmpty(node.LineColor))
            {
                sb.Append(" data-line-color=\"").Append(EscapeAttr(node.LineColor)).Append("\"");
            }

            if (node.LineWidthPt.HasValue)
            {
                sb.Append(" data-line-width=\"")
                    .Append(PptHtmlStyleIo.FormatPt(node.LineWidthPt.Value))
                    .Append("\"");
            }

            if (!string.IsNullOrEmpty(node.Align))
            {
                sb.Append(" data-align=\"").Append(EscapeAttr(node.Align)).Append("\"");
            }

            if (!string.IsNullOrEmpty(node.Valign))
            {
                sb.Append(" data-valign=\"").Append(EscapeAttr(node.Valign)).Append("\"");
            }

            if (node.TextWidthPct.HasValue)
            {
                sb.Append(" data-text-width=\"")
                    .Append(PptHtmlTextWidthIo.FormatPct(node.TextWidthPct.Value))
                    .Append("\"");
            }

            if (!string.IsNullOrEmpty(node.LineSpacing))
            {
                sb.Append(" data-line-spacing=\"").Append(EscapeAttr(node.LineSpacing)).Append("\"");
            }

            if (node.SpaceBeforePt.HasValue)
            {
                sb.Append(" data-space-before=\"")
                    .Append(PptHtmlParagraphIo.FormatPt(node.SpaceBeforePt.Value))
                    .Append("\"");
            }

            if (node.SpaceAfterPt.HasValue)
            {
                sb.Append(" data-space-after=\"")
                    .Append(PptHtmlParagraphIo.FormatPt(node.SpaceAfterPt.Value))
                    .Append("\"");
            }

            if (node.IndentLeftPt.HasValue)
            {
                sb.Append(" data-indent-left=\"")
                    .Append(PptHtmlParagraphIo.FormatPt(node.IndentLeftPt.Value))
                    .Append("\"");
            }

            if (node.IndentFirstPt.HasValue)
            {
                sb.Append(" data-indent-first=\"")
                    .Append(PptHtmlParagraphIo.FormatPt(node.IndentFirstPt.Value))
                    .Append("\"");
            }

            if (!string.IsNullOrEmpty(node.Bullet))
            {
                sb.Append(" data-bullet=\"").Append(EscapeAttr(node.Bullet)).Append("\"");
            }

            if (node.MarginLeftPt.HasValue)
            {
                sb.Append(" data-margin-left=\"")
                    .Append(PptHtmlStyleIo.FormatPt(node.MarginLeftPt.Value))
                    .Append("\"");
            }

            if (node.MarginRightPt.HasValue)
            {
                sb.Append(" data-margin-right=\"")
                    .Append(PptHtmlStyleIo.FormatPt(node.MarginRightPt.Value))
                    .Append("\"");
            }

            if (node.MarginTopPt.HasValue)
            {
                sb.Append(" data-margin-top=\"")
                    .Append(PptHtmlStyleIo.FormatPt(node.MarginTopPt.Value))
                    .Append("\"");
            }

            if (node.MarginBottomPt.HasValue)
            {
                sb.Append(" data-margin-bottom=\"")
                    .Append(PptHtmlStyleIo.FormatPt(node.MarginBottomPt.Value))
                    .Append("\"");
            }

            if (!string.IsNullOrEmpty(node.Style))
            {
                sb.Append(" style=\"").Append(EscapeAttr(node.Style)).Append("\"");
            }

            if (!string.IsNullOrEmpty(node.Name))
            {
                sb.Append(" data-name=\"").Append(EscapeAttr(node.Name)).Append("\"");
            }

            if (node.Rotation.HasValue && Math.Abs(node.Rotation.Value) > 0.05)
            {
                sb.Append(" data-rotation=\"")
                    .Append(node.Rotation.Value.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append("\"");
            }

            if (node.TextTruncated)
            {
                sb.Append(" data-truncated=\"true\"");
            }

            if (node.DepthCapped)
            {
                sb.Append(" data-depth-capped=\"true\"");
            }
        }

        private static void AppendSkeletonPreface(StringBuilder sb, PptHtmlReadResult result)
        {
            if (result.IsContentOnly)
            {
                string area = (result.ContentMinArea ?? DefaultContentMinArea)
                    .ToString("0.####", CultureInfo.InvariantCulture);
                sb.AppendLine(
                    "内容骨架：只留有字 / textbox / 占位符 / chart / table / media，以及页面积≥"
                    + area
                    + " 的真图；线、三角、菱形、圆点等无字装饰已去掉。");
                sb.AppendLine("要对齐装饰框看完整骨架；只要去掉栅格装饰用 compact=true。");
            }
            else if (result.IsCompact)
            {
                sb.AppendLine(
                    "精简骨架：已去掉带 data-rasterized-from 的栅格装饰（freeform 等），空组已去掉；这些装饰不占 200 顶。");
                sb.AppendLine("只要内容、不要矢量装饰用 content_only=true。要对齐装饰框看完整骨架。");
            }

            sb.AppendLine("骨架：本窗只展开 3 层（根不算层，从孩子起数；max_depth=3）。");
            List<string> ids = result.DepthCappedShapeIds;
            if (ids != null && ids.Count > 0)
            {
                sb.AppendLine("下列 group 还有更深子节点，本窗未写出。要继续展开请再调 F_read_ppt_html，传入该 shape_id：");
                for (int i = 0; i < ids.Count; i++)
                {
                    if (!string.IsNullOrEmpty(ids[i]))
                    {
                        sb.Append("  ").AppendLine(ids[i]);
                    }
                }

                return;
            }

            sb.AppendLine("本窗已写完，没有因深度截断的组。");
        }

        private static void CollectDepthCappedIds(IList<PptHtmlShapeNode> nodes, List<string> ids)
        {
            if (nodes == null)
            {
                return;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode node = nodes[i];
                if (node == null)
                {
                    continue;
                }

                if (node.DepthCapped && !string.IsNullOrEmpty(node.ShapeId))
                {
                    ids.Add(node.ShapeId);
                }

                CollectDepthCappedIds(node.Children, ids);
            }
        }

        private static string FormatPct(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return "0";
            }

            if (value < 0)
            {
                value = 0;
            }

            return value.ToString("0.####", CultureInfo.InvariantCulture);
        }

        private static string EscapeAttr(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text
                .Replace("&", "&amp;")
                .Replace("\"", "&quot;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        private static string EscapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }
    }
}
