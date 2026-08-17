using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlPowerPointReader
    {
        private const int MsoGroup = 6;
        private const int MsoPlaceholder = 14;
        private const int PpPlaceholderBody = 2;
        private const int PpPlaceholderVerticalBody = 17;

        public static bool TryRead(
            PowerPoint.Presentation presentation,
            string slideId,
            string channelId,
            string kind,
            out PptHtmlReadResult result,
            out string error)
        {
            result = null;
            error = null;
            if (presentation == null)
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (string.IsNullOrWhiteSpace(slideId))
            {
                error = "必须提供 slide_id（先 F_get_presentation_content）";
                return false;
            }

            string trimmed = slideId.Trim();
            if (!int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out int slideIdInt))
            {
                error = "slide_id 须为 SlideID 数字字符串，不要用第 N 页";
                return false;
            }

            PowerPoint.Slide slide;
            try
            {
                slide = FindSlideById(presentation, slideIdInt);
            }
            catch (Exception ex)
            {
                error = "查找幻灯片失败: " + ex.Message;
                return false;
            }

            if (slide == null)
            {
                error = "幻灯片不存在: slide_id=" + trimmed;
                return false;
            }

            string name;
            float slideWidth;
            float slideHeight;
            try
            {
                name = presentation.Name ?? "";
                slideWidth = presentation.PageSetup.SlideWidth;
                slideHeight = presentation.PageSetup.SlideHeight;
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            if (slideWidth <= 0 || slideHeight <= 0)
            {
                slideWidth = 720;
                slideHeight = 540;
            }

            var shapes = new List<PptHtmlShapeNode>();
            bool truncated = false;
            string truncatedReason = null;
            try
            {
                CollectShapes(
                    slide.Shapes,
                    trimmed,
                    slideWidth,
                    slideHeight,
                    shapes,
                    ref truncated,
                    ref truncatedReason);
                if (truncated && string.IsNullOrEmpty(truncatedReason))
                {
                    truncatedReason = "部分形状文本超过 " + PptHtmlReadResult.MaxTextChars + " 字符，已截断";
                }
            }
            catch (Exception ex)
            {
                error = "读取形状失败: " + ex.Message;
                return false;
            }

            result = new PptHtmlReadResult
            {
                ChannelId = channelId,
                Kind = kind,
                Name = name,
                SlideId = trimmed,
                Index = TryGetIndex(slide),
                Layout = TryGetLayout(slide),
                Hidden = TryGetHidden(slide),
                HasNotes = TryDetectHasNotes(slide),
                Shapes = shapes,
                Truncated = truncated,
                TruncatedReason = truncatedReason,
                ShapeCount = shapes.Count
            };
            return true;
        }

        private static PowerPoint.Slide FindSlideById(PowerPoint.Presentation presentation, int slideId)
        {
            try
            {
                return presentation.Slides.FindBySlideID(slideId);
            }
            catch (Exception)
            {
            }

            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                try
                {
                    if (slide.SlideID == slideId)
                    {
                        return slide;
                    }
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static void CollectShapes(
            PowerPoint.Shapes shapes,
            string slideId,
            float slideWidth,
            float slideHeight,
            List<PptHtmlShapeNode> output,
            ref bool truncated,
            ref string truncatedReason)
        {
            if (shapes == null)
            {
                return;
            }

            int count;
            try
            {
                count = shapes.Count;
            }
            catch (Exception)
            {
                return;
            }

            for (int i = 1; i <= count; i++)
            {
                if (output.Count >= PptHtmlReadResult.MaxShapes)
                {
                    truncated = true;
                    truncatedReason = "单页形状超过 " + PptHtmlReadResult.MaxShapes + "，已截断";
                    return;
                }

                PowerPoint.Shape shape;
                try
                {
                    shape = shapes[i];
                }
                catch (Exception)
                {
                    continue;
                }

                int shapeType = 0;
                try
                {
                    shapeType = (int)shape.Type;
                }
                catch (Exception)
                {
                }

                if (shapeType == MsoGroup)
                {
                    try
                    {
                        CollectGroupItems(
                            shape,
                            slideId,
                            slideWidth,
                            slideHeight,
                            output,
                            ref truncated,
                            ref truncatedReason);
                    }
                    catch (Exception)
                    {
                        AppendNode(shape, slideId, slideWidth, slideHeight, output, ref truncated);
                    }

                    if (truncated && output.Count >= PptHtmlReadResult.MaxShapes)
                    {
                        return;
                    }

                    continue;
                }

                AppendNode(shape, slideId, slideWidth, slideHeight, output, ref truncated);
            }
        }

        private static void CollectGroupItems(
            PowerPoint.Shape group,
            string slideId,
            float slideWidth,
            float slideHeight,
            List<PptHtmlShapeNode> output,
            ref bool truncated,
            ref string truncatedReason)
        {
            PowerPoint.GroupShapes items = group.GroupItems;
            int count = items.Count;
            for (int i = 1; i <= count; i++)
            {
                if (output.Count >= PptHtmlReadResult.MaxShapes)
                {
                    truncated = true;
                    truncatedReason = "单页形状超过 " + PptHtmlReadResult.MaxShapes + "，已截断";
                    return;
                }

                PowerPoint.Shape child = items[i];
                int childType = 0;
                try
                {
                    childType = (int)child.Type;
                }
                catch (Exception)
                {
                }

                if (childType == MsoGroup)
                {
                    CollectGroupItems(
                        child,
                        slideId,
                        slideWidth,
                        slideHeight,
                        output,
                        ref truncated,
                        ref truncatedReason);
                }
                else
                {
                    AppendNode(child, slideId, slideWidth, slideHeight, output, ref truncated);
                }
            }
        }

        private static void AppendNode(
            PowerPoint.Shape shape,
            string slideId,
            float slideWidth,
            float slideHeight,
            List<PptHtmlShapeNode> output,
            ref bool pageTextTruncated)
        {
            if (output.Count >= PptHtmlReadResult.MaxShapes)
            {
                return;
            }

            int comId;
            try
            {
                comId = shape.Id;
            }
            catch (Exception)
            {
                return;
            }

            int shapeType = 0;
            int? autoType = null;
            int? placeholderType = null;
            try
            {
                shapeType = (int)shape.Type;
            }
            catch (Exception)
            {
            }

            try
            {
                if (shapeType == MsoPlaceholder)
                {
                    placeholderType = Convert.ToInt32(shape.PlaceholderFormat.Type);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                autoType = Convert.ToInt32(shape.AutoShapeType);
            }
            catch (Exception)
            {
            }

            string typeName = PptShapeTypeMap.FromShapeType(shapeType, autoType, placeholderType);
            bool hasTable = false;
            try
            {
                hasTable = shape.HasTable == Office.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
            }

            if (hasTable)
            {
                typeName = "table";
            }

            string text = "";
            bool textTruncated = false;
            string innerHtml = null;
            if (typeName == "table")
            {
                innerHtml = TryBuildTableInner(shape, out textTruncated);
            }
            else if (typeName != "picture" && !PptShapeTypeMap.IsNonEditable(typeName))
            {
                text = TryReadText(shape);
                text = PptConventionHtml.TruncateText(text, out textTruncated);
            }

            if (textTruncated)
            {
                pageTextTruncated = true;
            }

            string style = TryBuildStyle(shape, slideWidth, slideHeight);
            double? rotation = null;
            try
            {
                rotation = shape.Rotation;
            }
            catch (Exception)
            {
            }

            string name = null;
            try
            {
                name = shape.Name;
            }
            catch (Exception)
            {
            }

            bool editable = !PptShapeTypeMap.IsNonEditable(typeName);
            string tag = PptShapeTypeMap.PreferTag(typeName, !string.IsNullOrEmpty(text));

            output.Add(new PptHtmlShapeNode
            {
                ShapeId = "sid" + slideId + "-s" + comId.ToString(CultureInfo.InvariantCulture),
                ShapeType = typeName,
                Tag = tag,
                Style = style,
                Text = text,
                InnerHtml = innerHtml,
                Editable = editable,
                Name = typeName == "picture" ? name : null,
                Rotation = rotation,
                TextTruncated = textTruncated
            });
        }

        private static string TryBuildStyle(PowerPoint.Shape shape, float slideWidth, float slideHeight)
        {
            try
            {
                float left = shape.Left;
                float top = shape.Top;
                float width = shape.Width;
                float height = shape.Height;
                return PptConventionHtml.BuildStyle(
                    left / slideWidth * 100.0,
                    top / slideHeight * 100.0,
                    width / slideWidth * 100.0,
                    height / slideHeight * 100.0);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadText(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return "";
                }

                return shape.TextFrame.TextRange.Text ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string TryBuildTableInner(PowerPoint.Shape shape, out bool textTruncated)
        {
            textTruncated = false;
            try
            {
                PowerPoint.Table table = shape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                var sb = new StringBuilder();
                for (int r = 1; r <= rows; r++)
                {
                    sb.Append("    <tr>");
                    for (int c = 1; c <= cols; c++)
                    {
                        string cellText = "";
                        try
                        {
                            cellText = table.Cell(r, c).Shape.TextFrame.TextRange.Text ?? "";
                        }
                        catch (Exception)
                        {
                        }

                        cellText = PptConventionHtml.TruncateText(cellText, out bool cellTrunc);
                        if (cellTrunc)
                        {
                            textTruncated = true;
                        }

                        sb.Append("<td>")
                            .Append(EscapeText(cellText))
                            .Append("</td>");
                    }

                    sb.AppendLine("</tr>");
                }

                return sb.ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string EscapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        private static int TryGetIndex(PowerPoint.Slide slide)
        {
            try
            {
                return slide.SlideIndex;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static string TryGetLayout(PowerPoint.Slide slide)
        {
            try
            {
                return slide.CustomLayout != null ? (slide.CustomLayout.Name ?? "") : "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool TryGetHidden(PowerPoint.Slide slide)
        {
            try
            {
                return slide.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool? TryDetectHasNotes(PowerPoint.Slide slide)
        {
            try
            {
                PowerPoint.SlideRange notes = slide.NotesPage;
                if (notes == null)
                {
                    return false;
                }

                foreach (PowerPoint.Shape shape in notes.Shapes)
                {
                    try
                    {
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            int ph = Convert.ToInt32(shape.PlaceholderFormat.Type);
                            if (ph != PpPlaceholderBody && ph != PpPlaceholderVerticalBody)
                            {
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }

                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                        {
                            continue;
                        }

                        string text = shape.TextFrame.TextRange.Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
