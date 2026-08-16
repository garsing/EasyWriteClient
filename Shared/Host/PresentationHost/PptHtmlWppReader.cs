using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlWppReader
    {
        private const int MsoGroup = 6;
        private const int MsoPlaceholder = 14;
        private const int PpPlaceholderBody = 2;
        private const int PpPlaceholderVerticalBody = 17;

        public static bool TryRead(
            object presentation,
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

            object slidesCollection;
            object slide;
            try
            {
                slidesCollection = WppCom.GetProperty(presentation, "Slides");
                slide = FindSlideById(slidesCollection, slideIdInt);
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
            double slideWidth;
            double slideHeight;
            try
            {
                name = WppCom.TryReadName(presentation) ?? "";
                object pageSetup = WppCom.GetProperty(presentation, "PageSetup");
                slideWidth = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideWidth"));
                slideHeight = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideHeight"));
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
                object shapesObj = WppCom.GetProperty(slide, "Shapes");
                CollectShapes(
                    shapesObj,
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

        private static object FindSlideById(object slidesCollection, int slideId)
        {
            if (slidesCollection == null)
            {
                return null;
            }

            try
            {
                // FindBySlideID(slideId)
                object found = slidesCollection.GetType().InvokeMember(
                    "FindBySlideID",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    slidesCollection,
                    new object[] { slideId });
                if (found != null)
                {
                    return found;
                }
            }
            catch (Exception)
            {
            }

            int count = Convert.ToInt32(WppCom.GetProperty(slidesCollection, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object slide = WppCom.GetIndexed(slidesCollection, i);
                if (slide == null)
                {
                    continue;
                }

                try
                {
                    object id = WppCom.GetProperty(slide, "SlideID");
                    if (id != null && Convert.ToInt32(id) == slideId)
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
            object shapes,
            string slideId,
            double slideWidth,
            double slideHeight,
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
                count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
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

                object shape = WppCom.GetIndexed(shapes, i);
                if (shape == null)
                {
                    continue;
                }

                int shapeType = TryGetInt(shape, "Type");
                if (shapeType == MsoGroup)
                {
                    try
                    {
                        object groupItems = WppCom.GetProperty(shape, "GroupItems");
                        CollectGroupItems(
                            groupItems,
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

                    continue;
                }

                AppendNode(shape, slideId, slideWidth, slideHeight, output, ref truncated);
            }
        }

        private static void CollectGroupItems(
            object groupItems,
            string slideId,
            double slideWidth,
            double slideHeight,
            List<PptHtmlShapeNode> output,
            ref bool truncated,
            ref string truncatedReason)
        {
            if (groupItems == null)
            {
                return;
            }

            int count = Convert.ToInt32(WppCom.GetProperty(groupItems, "Count"));
            for (int i = 1; i <= count; i++)
            {
                if (output.Count >= PptHtmlReadResult.MaxShapes)
                {
                    truncated = true;
                    truncatedReason = "单页形状超过 " + PptHtmlReadResult.MaxShapes + "，已截断";
                    return;
                }

                object child = WppCom.GetIndexed(groupItems, i);
                if (child == null)
                {
                    continue;
                }

                int childType = TryGetInt(child, "Type");
                if (childType == MsoGroup)
                {
                    object nested = WppCom.GetProperty(child, "GroupItems");
                    CollectGroupItems(
                        nested,
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
            object shape,
            string slideId,
            double slideWidth,
            double slideHeight,
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
                comId = Convert.ToInt32(WppCom.GetProperty(shape, "Id"));
            }
            catch (Exception)
            {
                return;
            }

            int shapeType = TryGetInt(shape, "Type");
            int? autoType = TryGetNullableInt(shape, "AutoShapeType");
            int? placeholderType = null;
            if (shapeType == MsoPlaceholder)
            {
                try
                {
                    object ph = WppCom.GetProperty(shape, "PlaceholderFormat");
                    placeholderType = Convert.ToInt32(WppCom.GetProperty(ph, "Type"));
                }
                catch (Exception)
                {
                }
            }

            string typeName = PptShapeTypeMap.FromShapeType(shapeType, autoType, placeholderType);
            if (IsTruthy(WppCom.GetProperty(shape, "HasTable")))
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
            else if (!PptShapeTypeMap.IsNonEditable(typeName))
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
                object rot = WppCom.GetProperty(shape, "Rotation");
                if (rot != null)
                {
                    rotation = Convert.ToDouble(rot);
                }
            }
            catch (Exception)
            {
            }

            string name = null;
            try
            {
                object n = WppCom.GetProperty(shape, "Name");
                name = n == null ? null : Convert.ToString(n);
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

        private static string TryBuildStyle(object shape, double slideWidth, double slideHeight)
        {
            try
            {
                double left = Convert.ToDouble(WppCom.GetProperty(shape, "Left"));
                double top = Convert.ToDouble(WppCom.GetProperty(shape, "Top"));
                double width = Convert.ToDouble(WppCom.GetProperty(shape, "Width"));
                double height = Convert.ToDouble(WppCom.GetProperty(shape, "Height"));
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

        private static string TryReadText(object shape)
        {
            try
            {
                if (!IsTruthy(WppCom.GetProperty(shape, "HasTextFrame")))
                {
                    return "";
                }

                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                object text = WppCom.GetProperty(tr, "Text");
                return text == null ? "" : Convert.ToString(text) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string TryBuildTableInner(object shape, out bool textTruncated)
        {
            textTruncated = false;
            try
            {
                object table = WppCom.GetProperty(shape, "Table");
                int rows = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(table, "Rows"), "Count"));
                int cols = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(table, "Columns"), "Count"));
                var sb = new StringBuilder();
                for (int r = 1; r <= rows; r++)
                {
                    sb.Append("    <tr>");
                    for (int c = 1; c <= cols; c++)
                    {
                        string cellText = "";
                        try
                        {
                            object cell = table.GetType().InvokeMember(
                                "Cell",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null,
                                table,
                                new object[] { r, c });
                            object cellShape = WppCom.GetProperty(cell, "Shape");
                            cellText = TryReadText(cellShape);
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

        private static int TryGetIndex(object slide)
        {
            try
            {
                return Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex"));
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static string TryGetLayout(object slide)
        {
            try
            {
                object layout = WppCom.GetProperty(slide, "CustomLayout");
                object name = WppCom.GetProperty(layout, "Name");
                return name == null ? "" : Convert.ToString(name) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool TryGetHidden(object slide)
        {
            try
            {
                object transition = WppCom.GetProperty(slide, "SlideShowTransition");
                return IsTruthy(WppCom.GetProperty(transition, "Hidden"));
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool? TryDetectHasNotes(object slide)
        {
            try
            {
                object notes = WppCom.GetProperty(slide, "NotesPage");
                if (notes == null)
                {
                    return false;
                }

                object shapes = WppCom.GetProperty(notes, "Shapes");
                int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object shape = WppCom.GetIndexed(shapes, i);
                    if (shape == null)
                    {
                        continue;
                    }

                    try
                    {
                        object ph = WppCom.GetProperty(shape, "PlaceholderFormat");
                        if (ph == null)
                        {
                            continue;
                        }

                        int phType = Convert.ToInt32(WppCom.GetProperty(ph, "Type"));
                        if (phType != PpPlaceholderBody && phType != PpPlaceholderVerticalBody)
                        {
                            continue;
                        }

                        string text = TryReadText(shape);
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

        private static int TryGetInt(object target, string name)
        {
            try
            {
                object v = WppCom.GetProperty(target, name);
                return v == null ? 0 : Convert.ToInt32(v);
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static int? TryGetNullableInt(object target, string name)
        {
            try
            {
                object v = WppCom.GetProperty(target, name);
                if (v == null)
                {
                    return null;
                }

                return Convert.ToInt32(v);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool IsTruthy(object value)
        {
            if (value == null)
            {
                return false;
            }

            try
            {
                if (value is bool b)
                {
                    return b;
                }

                int n = Convert.ToInt32(value);
                return n == -1 || n == 1;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
