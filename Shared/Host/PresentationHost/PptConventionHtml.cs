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
            sb.Append(BuildSectionHtml(result));
            return sb.ToString().TrimEnd();
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

        public static string BuildStyle(double leftPct, double topPct, double widthPct, double heightPct)
        {
            return "left:" + FormatPct(leftPct)
                + "%;top:" + FormatPct(topPct)
                + "%;width:" + FormatPct(widthPct)
                + "%;height:" + FormatPct(heightPct) + "%";
        }

        public static string TruncateText(string text, out bool truncated)
        {
            truncated = false;
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            if (normalized.Length <= PptHtmlReadResult.MaxTextChars)
            {
                return normalized;
            }

            truncated = true;
            return normalized.Substring(0, PptHtmlReadResult.MaxTextChars);
        }

        private static void AppendShape(StringBuilder sb, PptHtmlShapeNode node)
        {
            string tag = string.IsNullOrEmpty(node.Tag) ? "div" : node.Tag;
            if (tag == "img")
            {
                sb.Append("  <img ShapeId=\"").Append(EscapeAttr(node.ShapeId)).Append("\"");
                AppendCommonAttrs(sb, node);
                sb.AppendLine(" />");
                return;
            }

            if (tag == "table")
            {
                sb.Append("  <table ShapeId=\"").Append(EscapeAttr(node.ShapeId)).Append("\"");
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

                sb.AppendLine("  </table>");
                return;
            }

            if (node.ShapeType == "chart")
            {
                sb.Append("  <div ShapeId=\"").Append(EscapeAttr(node.ShapeId)).Append("\"");
                AppendCommonAttrs(sb, node);
                PptHtmlChartIo.AppendFormatAttrs(sb, node.ChartFormat);
                sb.AppendLine(">");
                sb.AppendLine("    <table>");
                if (!string.IsNullOrEmpty(node.InnerHtml))
                {
                    sb.Append(node.InnerHtml);
                    if (!node.InnerHtml.EndsWith("\n", StringComparison.Ordinal))
                    {
                        sb.AppendLine();
                    }
                }

                sb.AppendLine("    </table>");
                sb.AppendLine("  </div>");
                return;
            }

            sb.Append("  <").Append(tag)
                .Append(" ShapeId=\"").Append(EscapeAttr(node.ShapeId)).Append("\"");
            AppendCommonAttrs(sb, node);
            sb.Append(">");
            sb.Append(EscapeText(node.Text));
            sb.Append("</").Append(tag).AppendLine(">");
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
