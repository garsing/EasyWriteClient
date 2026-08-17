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

                // B2：整组栅格为一张 picture，不再展开子项
                AppendNode(shape, slideId, slideWidth, slideHeight, output, ref truncated);
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

            string rasterizedFrom = null;
            if (PptShapeTypeMap.ShouldRasterizeAsPicture(typeName))
            {
                rasterizedFrom = typeName;
                typeName = "picture";
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
            string fill = null;
            string fontColor = null;
            double? fontSize = null;
            bool? fontBold = null;
            string fontName = null;
            string lineColor = null;
            double? lineWidth = null;
            if (typeName != "picture" && typeName != "media")
            {
                fill = TryReadFill(shape);
                if (typeName != "table")
                {
                    fontColor = TryReadFontColor(shape);
                    fontSize = TryReadFontSize(shape);
                    fontBold = TryReadFontBold(shape);
                    fontName = TryReadFontName(shape);
                }

                lineColor = TryReadLineColor(shape);
                lineWidth = TryReadLineWidth(shape);
            }

            int? z = TryReadZ(shape);

            output.Add(new PptHtmlShapeNode
            {
                ShapeId = "sid" + slideId + "-s" + comId.ToString(CultureInfo.InvariantCulture),
                ShapeType = typeName,
                Tag = tag,
                Style = style,
                Text = text,
                InnerHtml = innerHtml,
                Editable = editable,
                Fill = fill,
                FontColor = fontColor,
                FontSizePt = fontSize,
                FontBold = fontBold,
                FontName = fontName,
                Z = z,
                LineColor = lineColor,
                LineWidthPt = lineWidth,
                RasterizedFrom = rasterizedFrom,
                Name = typeName == "picture" ? name : null,
                Rotation = rotation,
                TextTruncated = textTruncated
            });
        }

        private static double? TryReadFontSize(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return null;
                }

                float size = shape.TextFrame.TextRange.Font.Size;
                if (size <= 0)
                {
                    return null;
                }

                return size;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool? TryReadFontBold(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return null;
                }

                Office.MsoTriState bold = shape.TextFrame.TextRange.Font.Bold;
                if (bold == Office.MsoTriState.msoTriStateMixed)
                {
                    return null;
                }

                return bold == Office.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadFontName(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return null;
                }

                PowerPoint.Font font = shape.TextFrame.TextRange.Font;
                // 中文正文看 NameFarEast；仅设 Name 时常仍显示「等线」
                string name = null;
                try
                {
                    name = font.NameFarEast;
                }
                catch (Exception)
                {
                }

                if (string.IsNullOrWhiteSpace(name))
                {
                    try
                    {
                        name = font.Name;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (string.IsNullOrWhiteSpace(name))
                {
                    try
                    {
                        name = font.NameAscii;
                    }
                    catch (Exception)
                    {
                    }
                }

                return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int? TryReadZ(PowerPoint.Shape shape)
        {
            try
            {
                int z = shape.ZOrderPosition;
                if (z < 0)
                {
                    return null;
                }

                return z;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadLineColor(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                {
                    return "none";
                }

                int rgb = shape.Line.ForeColor.RGB;
                return PptHtmlStyleIo.FormatOfficeRgb(rgb);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static double? TryReadLineWidth(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                {
                    return null;
                }

                float w = shape.Line.Weight;
                if (w < 0)
                {
                    return null;
                }

                return w;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadFill(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.Fill.Visible == Office.MsoTriState.msoFalse)
                {
                    return "none";
                }

                int rgb = shape.Fill.ForeColor.RGB;
                return PptHtmlStyleIo.FormatOfficeRgb(rgb);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadFontColor(PowerPoint.Shape shape)
        {
            try
            {
                // TextFrame2 优先；主题色 RGB 常为 0/负值，须按 ObjectThemeColor 查色表。
                try
                {
                    var fore = shape.TextFrame2.TextRange.Font.Fill.ForeColor;
                    int type2 = (int)fore.Type;
                    int rgb2 = fore.RGB;
                    int rgbNorm = rgb2 & 0x00FFFFFF;
                    Office.MsoThemeColorIndex themeIdx = Office.MsoThemeColorIndex.msoNotThemeColor;
                    float brightness = 0f;
                    try
                    {
                        themeIdx = fore.ObjectThemeColor;
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        brightness = fore.Brightness;
                    }
                    catch (Exception)
                    {
                    }

                    LogFontColorRead(
                        shape,
                        "tf2-raw",
                        type2,
                        rgb2,
                        rgbNorm != 0 ? PptHtmlStyleIo.FormatOfficeRgb(rgbNorm) : null,
                        themeIdx,
                        brightness);

                    // 有明确非零 RGB（含负 Long 的低 24 位）且未挂主题 → 直接用
                    if (rgbNorm != 0 && themeIdx == Office.MsoThemeColorIndex.msoNotThemeColor)
                    {
                        string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                        LogFontColorRead(shape, "tf2", type2, rgbNorm, hex, themeIdx, brightness);
                        return hex;
                    }

                    // 主题色（或 RGB=0/可疑）：从 ThemeColorScheme 还原
                    if (themeIdx != Office.MsoThemeColorIndex.msoNotThemeColor
                        || rgbNorm == 0)
                    {
                        if (TryResolveThemeFontRgb(shape, fore, out int themeRgb))
                        {
                            int tNorm = themeRgb & 0x00FFFFFF;
                            if (tNorm != 0)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(tNorm);
                                LogFontColorRead(shape, "theme", type2, tNorm, hex, themeIdx, brightness);
                                return hex;
                            }
                        }
                    }

                    if (rgbNorm != 0)
                    {
                        string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                        LogFontColorRead(shape, "tf2-fallback", type2, rgbNorm, hex, themeIdx, brightness);
                        return hex;
                    }

                    // 挂了主题却解不出：宁可不写 data-font-color，也不要误导出 #000000
                    if (themeIdx != Office.MsoThemeColorIndex.msoNotThemeColor)
                    {
                        LogFontColorRead(shape, "theme-omit", type2, 0, null, themeIdx, brightness);
                        return null;
                    }

                    // 纯 RGB 且为 0 → 真黑
                    if (type2 == 1 /* msoColorTypeRGB */)
                    {
                        string hex = PptHtmlStyleIo.FormatOfficeRgb(0);
                        LogFontColorRead(shape, "tf2-black", type2, 0, hex, themeIdx, brightness);
                        return hex;
                    }

                    LogFontColorRead(shape, "tf2-skip0", type2, rgb2, null, themeIdx, brightness);
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.PptHtml, "font_color tf2 err: " + ex.Message);
                }

                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return null;
                }

                var color = shape.TextFrame.TextRange.Font.Color;
                int rgb = color.RGB;
                int type = (int)color.Type;
                int norm = rgb & 0x00FFFFFF;
                if (norm == 0 && type != 1)
                {
                    LogFontColorRead(shape, "tf1-skip0", type, rgb, null);
                    return null;
                }

                string hex1 = PptHtmlStyleIo.FormatOfficeRgb(norm);
                LogFontColorRead(shape, "tf1", type, norm, hex1);
                return hex1;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void LogFontColorRead(
            PowerPoint.Shape shape,
            string via,
            int type,
            int rgb,
            string hex,
            Office.MsoThemeColorIndex themeIdx = Office.MsoThemeColorIndex.msoNotThemeColor,
            float brightness = 0f)
        {
            if (!EasyWriteDiagnostics.IsEnabled(DebugCategory.PptHtml))
            {
                return;
            }

            string name = "";
            string text = "";
            try
            {
                name = shape.Name ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    text = shape.TextFrame.TextRange.Text ?? "";
                    if (text.Length > 24)
                    {
                        text = text.Substring(0, 24);
                    }

                    text = text.Replace("\r", " ").Replace("\n", " ").Trim();
                }
            }
            catch (Exception)
            {
            }

            EasyWriteDiagnostics.Log(
                DebugCategory.PptHtml,
                "font_color via=" + via
                + " type=" + type.ToString(CultureInfo.InvariantCulture)
                + " rgb=" + rgb.ToString(CultureInfo.InvariantCulture)
                + " hex=" + (hex ?? "null")
                + " theme=" + ((int)themeIdx).ToString(CultureInfo.InvariantCulture)
                + " bright=" + brightness.ToString("0.###", CultureInfo.InvariantCulture)
                + " name=" + name
                + " text=" + text);
        }

        /// <summary>主题色 RGB 常为 0：从 SlideMaster.ThemeColorScheme + Brightness 还原。</summary>
        private static bool TryResolveThemeFontRgb(
            PowerPoint.Shape shape,
            Office.ColorFormat fore,
            out int rgb)
        {
            rgb = 0;
            try
            {
                Office.MsoThemeColorIndex themeIdx = fore.ObjectThemeColor;
                if (themeIdx == Office.MsoThemeColorIndex.msoNotThemeColor)
                {
                    return false;
                }

                if (!PptHtmlStyleIo.TryMapThemeColorIndexToSchemeIndex((int)themeIdx, out int schemeInt))
                {
                    EasyWriteDiagnostics.Log(
                        DebugCategory.PptHtml,
                        "font_color theme map fail idx=" + ((int)themeIdx).ToString(CultureInfo.InvariantCulture));
                    return false;
                }

                Office.MsoThemeColorSchemeIndex schemeIdx =
                    (Office.MsoThemeColorSchemeIndex)schemeInt;

                Office.ThemeColorScheme scheme = TryGetThemeColorScheme(shape);
                if (scheme == null)
                {
                    return false;
                }

                Office.ThemeColor themeColor = scheme.Colors(schemeIdx);
                rgb = themeColor.RGB & 0x00FFFFFF;

                float brightness = 0f;
                try
                {
                    brightness = fore.Brightness;
                }
                catch (Exception)
                {
                }

                if (Math.Abs(brightness) > 0.0001f)
                {
                    rgb = ApplyThemeBrightness(rgb, brightness);
                }

                return true;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.PptHtml, "font_color theme resolve err: " + ex.Message);
                return false;
            }
        }

        private static Office.ThemeColorScheme TryGetThemeColorScheme(PowerPoint.Shape shape)
        {
            try
            {
                PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
                if (slide != null)
                {
                    try
                    {
                        return slide.Design.SlideMaster.Theme.ThemeColorScheme;
                    }
                    catch (Exception)
                    {
                    }

                    PowerPoint.Presentation pres = slide.Parent as PowerPoint.Presentation;
                    if (pres != null)
                    {
                        return pres.SlideMaster.Theme.ThemeColorScheme;
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static int ApplyThemeBrightness(int officeRgb, float brightness)
        {
            int v = officeRgb & 0x00FFFFFF;
            int r = v & 0xFF;
            int g = (v >> 8) & 0xFF;
            int b = (v >> 16) & 0xFF;
            if (brightness > 0f)
            {
                r = (int)(r + (255 - r) * brightness);
                g = (int)(g + (255 - g) * brightness);
                b = (int)(b + (255 - b) * brightness);
            }
            else
            {
                float f = 1f + brightness;
                r = (int)(r * f);
                g = (int)(g * f);
                b = (int)(b * f);
            }

            r = Math.Max(0, Math.Min(255, r));
            g = Math.Max(0, Math.Min(255, g));
            b = Math.Max(0, Math.Min(255, b));
            return r | (g << 8) | (b << 16);
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
