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
            // 字色诊断：始终写会话文件，便于定位主题色误读/省略
            var fontDbg = new PptHtmlReadDebug();
            fontDbg.Line(
                "begin slide_id=" + trimmed
                + " slideSize=" + slideWidth.ToString("0.#", CultureInfo.InvariantCulture)
                + "x" + slideHeight.ToString("0.#", CultureInfo.InvariantCulture));
            try
            {
                if (!CollectShapes(
                    slide.Shapes,
                    trimmed,
                    slideWidth,
                    slideHeight,
                    shapes,
                    ref truncated,
                    ref truncatedReason,
                    fontDbg,
                    out string collectError))
                {
                    error = collectError;
                    return false;
                }

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

            string debugFile = fontDbg.TryWriteToSession(trimmed, out string debugWriteErr);
            if (!string.IsNullOrEmpty(debugWriteErr))
            {
                fontDbg.Line("write_debug_file_fail: " + debugWriteErr);
            }
            else
            {
                fontDbg.Line("debug_file=" + (debugFile ?? ""));
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
                ShapeCount = shapes.Count,
                DebugFilename = debugFile,
                DebugTrace = new List<string>(fontDbg.Lines)
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

        private static bool CollectShapes(
            PowerPoint.Shapes shapes,
            string slideId,
            float slideWidth,
            float slideHeight,
            List<PptHtmlShapeNode> output,
            ref bool truncated,
            ref string truncatedReason,
            PptHtmlReadDebug fontDbg,
            out string error)
        {
            error = null;
            if (shapes == null)
            {
                return true;
            }

            int count;
            try
            {
                count = shapes.Count;
            }
            catch (Exception)
            {
                return true;
            }

            for (int i = 1; i <= count; i++)
            {
                if (output.Count >= PptHtmlReadResult.MaxShapes)
                {
                    truncated = true;
                    truncatedReason = "单页形状超过 " + PptHtmlReadResult.MaxShapes + "，已截断";
                    return true;
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
                if (!AppendNode(
                    shape,
                    slideId,
                    slideWidth,
                    slideHeight,
                    output,
                    ref truncated,
                    ref truncatedReason,
                    fontDbg,
                    out error))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool AppendNode(
            PowerPoint.Shape shape,
            string slideId,
            float slideWidth,
            float slideHeight,
            List<PptHtmlShapeNode> output,
            ref bool pageTextTruncated,
            ref string truncatedReason,
            PptHtmlReadDebug fontDbg,
            out string error)
        {
            error = null;
            if (output.Count >= PptHtmlReadResult.MaxShapes)
            {
                return true;
            }

            int comId;
            try
            {
                comId = shape.Id;
            }
            catch (Exception)
            {
                return true;
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
            else if (PptHtmlChartIo.LooksLikeChart(shape))
            {
                typeName = "chart";
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
            PptHtmlChartFormat chartFormat = null;
            if (typeName == "chart")
            {
                if (!PptHtmlChartIo.TryRead(shape, out PptHtmlChartReadModel model, out error))
                {
                    error = "无法读取图表 sid" + slideId + "-s"
                        + comId.ToString(CultureInfo.InvariantCulture) + ": " + (error ?? "");
                    return false;
                }

                innerHtml = PptHtmlChartIo.BuildInnerHtml(model.Grid);
                chartFormat = model.Format;
                if (model.Grid != null && model.Grid.Truncated)
                {
                    textTruncated = true;
                    truncatedReason = "图表数据超过 "
                        + PptHtmlChartIo.MaxRows + "×" + PptHtmlChartIo.MaxCols + "，已截断";
                }
            }
            else if (typeName == "table")
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
            string align = null;
            string lineSpacing = null;
            double? spaceBefore = null;
            double? spaceAfter = null;
            double? indentLeft = null;
            double? indentFirst = null;
            string bullet = null;
            double? marginLeft = null;
            double? marginRight = null;
            double? marginTop = null;
            double? marginBottom = null;
            if (typeName != "picture" && typeName != "media" && typeName != "chart")
            {
                fill = TryReadFill(shape);
                if (typeName != "table")
                {
                    fontColor = TryReadFontColor(shape, fontDbg);
                    fontSize = TryReadFontSize(shape);
                    fontBold = TryReadFontBold(shape);
                    fontName = TryReadFontName(shape);
                    PptHtmlParagraphIo.Snapshot para = PptHtmlParagraphIo.TryReadFromShape(shape);
                    if (para != null)
                    {
                        align = para.Align;
                        lineSpacing = para.LineSpacing;
                        spaceBefore = para.SpaceBeforePt;
                        spaceAfter = para.SpaceAfterPt;
                        indentLeft = para.IndentLeftPt;
                        indentFirst = para.IndentFirstPt;
                        bullet = para.Bullet;
                    }

                    TryReadTextMargins(
                        shape,
                        out marginLeft,
                        out marginRight,
                        out marginTop,
                        out marginBottom);
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
                Align = align,
                LineSpacing = lineSpacing,
                SpaceBeforePt = spaceBefore,
                SpaceAfterPt = spaceAfter,
                IndentLeftPt = indentLeft,
                IndentFirstPt = indentFirst,
                Bullet = bullet,
                MarginLeftPt = marginLeft,
                MarginRightPt = marginRight,
                MarginTopPt = marginTop,
                MarginBottomPt = marginBottom,
                RasterizedFrom = rasterizedFrom,
                Name = typeName == "picture" ? name : null,
                Rotation = rotation,
                TextTruncated = textTruncated,
                ChartFormat = chartFormat
            });
            return true;
        }

        private static void TryReadTextMargins(
            PowerPoint.Shape shape,
            out double? left,
            out double? right,
            out double? top,
            out double? bottom)
        {
            left = right = top = bottom = null;
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return;
                }

                PowerPoint.TextFrame tf = shape.TextFrame;
                left = tf.MarginLeft;
                right = tf.MarginRight;
                top = tf.MarginTop;
                bottom = tf.MarginBottom;
            }
            catch (Exception)
            {
            }
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

        private static string TryReadFontColor(PowerPoint.Shape shape, PptHtmlReadDebug dbg)
        {
            string tag = ShapeProbeTag(shape);
            try
            {
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
                    catch (Exception ex)
                    {
                        dbg?.Probe(tag + " tf2 ObjectThemeColor err=" + ex.Message);
                    }

                    try
                    {
                        brightness = fore.Brightness;
                    }
                    catch (Exception ex)
                    {
                        dbg?.Probe(tag + " tf2 Brightness err=" + ex.Message);
                    }

                    dbg?.Probe(
                        tag
                        + " tf2-raw type=" + type2.ToString(CultureInfo.InvariantCulture)
                        + " rgbRaw=" + rgb2.ToString(CultureInfo.InvariantCulture)
                        + " rgbNorm=" + rgbNorm.ToString(CultureInfo.InvariantCulture)
                        + " hexNorm=" + (rgbNorm != 0 ? PptHtmlStyleIo.FormatOfficeRgb(rgbNorm) : "0")
                        + " theme=" + ((int)themeIdx).ToString(CultureInfo.InvariantCulture)
                        + " bright=" + brightness.ToString("0.###", CultureInfo.InvariantCulture));

                    // Mixed：整段无统一色，改读首字符
                    if (IsMixedFontColorSignal(type2, (int)themeIdx, rgb2))
                    {
                        string mixedHex = TryReadFontColorFromFirstCharacter(shape, dbg, tag);
                        if (!string.IsNullOrEmpty(mixedHex))
                        {
                            return mixedHex;
                        }

                        dbg?.Probe(tag + " mixed-char miss -> continue");
                    }

                    bool hasTheme = IsResolvableThemeIndex(themeIdx);

                    if (rgbNorm != 0 && !hasTheme)
                    {
                        string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                        dbg?.Probe(tag + " RESULT via=tf2 hex=" + hex);
                        return hex;
                    }

                    if (hasTheme || rgbNorm == 0)
                    {
                        if (hasTheme
                            && TryResolveThemeFontRgb(shape, fore, dbg, tag, out int themeRgb, out string resolveNote))
                        {
                            int tNorm = themeRgb & 0x00FFFFFF;
                            dbg?.Probe(
                                tag
                                + " theme-resolve ok note=" + (resolveNote ?? "")
                                + " rgb=" + tNorm.ToString(CultureInfo.InvariantCulture)
                                + " hex=" + (tNorm != 0 ? PptHtmlStyleIo.FormatOfficeRgb(tNorm) : "#000000"));
                            if (tNorm != 0)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(tNorm);
                                dbg?.Probe(tag + " RESULT via=theme hex=" + hex);
                                return hex;
                            }

                            dbg?.Probe(tag + " theme-resolve rgb=0 -> continue");
                        }
                        else if (hasTheme)
                        {
                            dbg?.Probe(tag + " theme-resolve FAIL");
                        }
                    }

                    if (rgbNorm != 0)
                    {
                        string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                        dbg?.Probe(tag + " RESULT via=tf2-fallback hex=" + hex);
                        return hex;
                    }

                    if (hasTheme)
                    {
                        dbg?.Probe(tag + " RESULT via=theme-omit hex=null");
                        return null;
                    }

                    if (type2 == 1)
                    {
                        string hex = PptHtmlStyleIo.FormatOfficeRgb(0);
                        dbg?.Probe(tag + " RESULT via=tf2-black hex=" + hex);
                        return hex;
                    }

                    dbg?.Probe(tag + " tf2-skip0 -> try TextFrame / first-char");
                    string charHex = TryReadFontColorFromFirstCharacter(shape, dbg, tag);
                    if (!string.IsNullOrEmpty(charHex))
                    {
                        return charHex;
                    }
                }
                catch (Exception ex)
                {
                    dbg?.Probe(tag + " tf2 err=" + ex.Message);
                    string charHex = TryReadFontColorFromFirstCharacter(shape, dbg, tag);
                    if (!string.IsNullOrEmpty(charHex))
                    {
                        return charHex;
                    }
                }

                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    dbg?.Probe(tag + " RESULT via=no-textframe hex=null");
                    return null;
                }

                var color = shape.TextFrame.TextRange.Font.Color;
                int rgb = color.RGB;
                int type = (int)color.Type;
                int norm = rgb & 0x00FFFFFF;
                int scheme = -1;
                int theme1 = 0;
                float bright1 = 0f;
                try { scheme = (int)color.SchemeColor; } catch (Exception) { }
                try { theme1 = (int)color.ObjectThemeColor; } catch (Exception) { }
                try { bright1 = color.Brightness; } catch (Exception) { }

                dbg?.Probe(
                    tag
                    + " tf1-raw type=" + type.ToString(CultureInfo.InvariantCulture)
                    + " rgbRaw=" + rgb.ToString(CultureInfo.InvariantCulture)
                    + " rgbNorm=" + norm.ToString(CultureInfo.InvariantCulture)
                    + " scheme=" + scheme.ToString(CultureInfo.InvariantCulture)
                    + " theme=" + theme1.ToString(CultureInfo.InvariantCulture)
                    + " bright=" + bright1.ToString("0.###", CultureInfo.InvariantCulture));

                if (IsMixedFontColorSignal(type, theme1, rgb))
                {
                    string mixedHex = TryReadFontColorFromFirstCharacter(shape, dbg, tag);
                    if (!string.IsNullOrEmpty(mixedHex))
                    {
                        return mixedHex;
                    }
                }

                if (norm == 0 && type != 1)
                {
                    dbg?.Probe(tag + " RESULT via=tf1-skip0 hex=null");
                    return null;
                }

                string hex1 = PptHtmlStyleIo.FormatOfficeRgb(norm);
                dbg?.Probe(tag + " RESULT via=tf1 hex=" + hex1);
                return hex1;
            }
            catch (Exception ex)
            {
                dbg?.Probe(tag + " RESULT via=exception hex=null err=" + ex.Message);
                return null;
            }
        }

        /// <summary>msoColorTypeMixed / msoThemeColorMixed = -2；RGB 哨兵为 int.MinValue。</summary>
        private static bool IsMixedFontColorSignal(int type, int themeIdx, int rgbRaw)
        {
            return type == -2 || themeIdx == -2 || rgbRaw == int.MinValue;
        }

        private static bool IsResolvableThemeIndex(Office.MsoThemeColorIndex idx)
        {
            return (int)idx > 0;
        }

        /// <summary>Mixed 时整段无统一色：读首字符 ForeColor。</summary>
        private static string TryReadFontColorFromFirstCharacter(
            PowerPoint.Shape shape,
            PptHtmlReadDebug dbg,
            string tag)
        {
            try
            {
                Office.TextRange2 tr2 = shape.TextFrame2.TextRange;
                if (tr2 != null && tr2.Length >= 1)
                {
                    Office.TextRange2 ch = tr2.Characters[1, 1];
                    var fore = ch.Font.Fill.ForeColor;
                    int type = (int)fore.Type;
                    int rgbRaw = fore.RGB;
                    int rgbNorm = rgbRaw & 0x00FFFFFF;
                    Office.MsoThemeColorIndex themeIdx = Office.MsoThemeColorIndex.msoNotThemeColor;
                    float brightness = 0f;
                    try { themeIdx = fore.ObjectThemeColor; } catch (Exception) { }
                    try { brightness = fore.Brightness; } catch (Exception) { }

                    dbg?.Probe(
                        tag
                        + " char1-tf2 type=" + type.ToString(CultureInfo.InvariantCulture)
                        + " rgbRaw=" + rgbRaw.ToString(CultureInfo.InvariantCulture)
                        + " rgbNorm=" + rgbNorm.ToString(CultureInfo.InvariantCulture)
                        + " theme=" + ((int)themeIdx).ToString(CultureInfo.InvariantCulture)
                        + " bright=" + brightness.ToString("0.###", CultureInfo.InvariantCulture));

                    if (!IsMixedFontColorSignal(type, (int)themeIdx, rgbRaw))
                    {
                        if (rgbNorm != 0 && !IsResolvableThemeIndex(themeIdx))
                        {
                            string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                            dbg?.Probe(tag + " RESULT via=char1-tf2 hex=" + hex);
                            return hex;
                        }

                        if (IsResolvableThemeIndex(themeIdx)
                            && TryResolveThemeFontRgb(shape, fore, dbg, tag + "/char1", out int themeRgb, out string note))
                        {
                            int tNorm = themeRgb & 0x00FFFFFF;
                            dbg?.Probe(tag + " char1-theme note=" + (note ?? ""));
                            if (tNorm != 0)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(tNorm);
                                dbg?.Probe(tag + " RESULT via=char1-theme hex=" + hex);
                                return hex;
                            }
                        }

                        if (rgbNorm != 0)
                        {
                            string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                            dbg?.Probe(tag + " RESULT via=char1-tf2-fallback hex=" + hex);
                            return hex;
                        }

                        if (type == 1)
                        {
                            string hex = PptHtmlStyleIo.FormatOfficeRgb(0);
                            dbg?.Probe(tag + " RESULT via=char1-tf2-black hex=" + hex);
                            return hex;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbg?.Probe(tag + " char1-tf2 err=" + ex.Message);
            }

            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.TextRange tr = shape.TextFrame.TextRange;
                    if (tr != null && tr.Length >= 1)
                    {
                        PowerPoint.TextRange ch = tr.Characters(1, 1);
                        var color = ch.Font.Color;
                        int type = (int)color.Type;
                        int rgbRaw = color.RGB;
                        int rgbNorm = rgbRaw & 0x00FFFFFF;
                        int theme1 = 0;
                        float bright1 = 0f;
                        try { theme1 = (int)color.ObjectThemeColor; } catch (Exception) { }
                        try { bright1 = color.Brightness; } catch (Exception) { }

                        dbg?.Probe(
                            tag
                            + " char1-tf1 type=" + type.ToString(CultureInfo.InvariantCulture)
                            + " rgbRaw=" + rgbRaw.ToString(CultureInfo.InvariantCulture)
                            + " rgbNorm=" + rgbNorm.ToString(CultureInfo.InvariantCulture)
                            + " theme=" + theme1.ToString(CultureInfo.InvariantCulture)
                            + " bright=" + bright1.ToString("0.###", CultureInfo.InvariantCulture));

                        if (!IsMixedFontColorSignal(type, theme1, rgbRaw))
                        {
                            if (rgbNorm != 0 && theme1 <= 0)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                                dbg?.Probe(tag + " RESULT via=char1-tf1 hex=" + hex);
                                return hex;
                            }

                            if (theme1 > 0
                                && PptHtmlStyleIo.TryMapThemeColorIndexToSchemeIndex(theme1, out int schemeInt))
                            {
                                try
                                {
                                    Office.ThemeColorScheme scheme = TryGetThemeColorScheme(shape, out _);
                                    if (scheme != null)
                                    {
                                        int themeRgb = scheme.Colors((Office.MsoThemeColorSchemeIndex)schemeInt).RGB
                                            & 0x00FFFFFF;
                                        if (Math.Abs(bright1) > 0.0001f)
                                        {
                                            themeRgb = ApplyThemeBrightness(themeRgb, bright1);
                                        }

                                        if (themeRgb != 0)
                                        {
                                            string hex = PptHtmlStyleIo.FormatOfficeRgb(themeRgb);
                                            dbg?.Probe(tag + " RESULT via=char1-tf1-theme hex=" + hex);
                                            return hex;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    dbg?.Probe(tag + " char1-tf1-theme err=" + ex.Message);
                                }
                            }

                            if (rgbNorm != 0)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                                dbg?.Probe(tag + " RESULT via=char1-tf1-fallback hex=" + hex);
                                return hex;
                            }

                            if (type == 1)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(0);
                                dbg?.Probe(tag + " RESULT via=char1-tf1-black hex=" + hex);
                                return hex;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dbg?.Probe(tag + " char1-tf1 err=" + ex.Message);
            }

            return null;
        }

        private static string ShapeProbeTag(PowerPoint.Shape shape)
        {
            string id = "?";
            string name = "";
            string text = "";
            try { id = shape.Id.ToString(CultureInfo.InvariantCulture); } catch (Exception) { }
            try { name = shape.Name ?? ""; } catch (Exception) { }
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    text = shape.TextFrame.TextRange.Text ?? "";
                    text = text.Replace("\r", " ").Replace("\n", " ").Trim();
                    if (text.Length > 20) text = text.Substring(0, 20);
                }
            }
            catch (Exception) { }
            return "font id=" + id + " name=" + name + " text=[" + text + "]";
        }

        private static bool TryResolveThemeFontRgb(
            PowerPoint.Shape shape,
            Office.ColorFormat fore,
            PptHtmlReadDebug dbg,
            string tag,
            out int rgb,
            out string note)
        {
            rgb = 0;
            note = null;
            try
            {
                Office.MsoThemeColorIndex themeIdx = fore.ObjectThemeColor;
                if (themeIdx == Office.MsoThemeColorIndex.msoNotThemeColor)
                {
                    note = "themeIdx=0(notTheme)";
                    return false;
                }

                if (!PptHtmlStyleIo.TryMapThemeColorIndexToSchemeIndex((int)themeIdx, out int schemeInt))
                {
                    note = "mapFail themeIdx=" + ((int)themeIdx).ToString(CultureInfo.InvariantCulture);
                    return false;
                }

                Office.MsoThemeColorSchemeIndex schemeIdx = (Office.MsoThemeColorSchemeIndex)schemeInt;
                Office.ThemeColorScheme scheme = TryGetThemeColorScheme(shape, out string schemeSrc);
                if (scheme == null)
                {
                    note = "noThemeColorScheme src=" + (schemeSrc ?? "");
                    return false;
                }

                Office.ThemeColor themeColor = scheme.Colors(schemeIdx);
                int baseRgb = themeColor.RGB;
                rgb = baseRgb & 0x00FFFFFF;
                float brightness = 0f;
                try { brightness = fore.Brightness; } catch (Exception) { }
                int beforeBright = rgb;
                if (Math.Abs(brightness) > 0.0001f)
                {
                    rgb = ApplyThemeBrightness(rgb, brightness);
                }

                note = "src=" + (schemeSrc ?? "?")
                    + " themeIdx=" + ((int)themeIdx).ToString(CultureInfo.InvariantCulture)
                    + " schemeIdx=" + schemeInt.ToString(CultureInfo.InvariantCulture)
                    + " baseRgb=" + (baseRgb & 0x00FFFFFF).ToString(CultureInfo.InvariantCulture)
                    + " baseHex=" + PptHtmlStyleIo.FormatOfficeRgb(beforeBright)
                    + " bright=" + brightness.ToString("0.###", CultureInfo.InvariantCulture)
                    + " afterHex=" + PptHtmlStyleIo.FormatOfficeRgb(rgb);
                return true;
            }
            catch (Exception ex)
            {
                note = "ex=" + ex.Message;
                dbg?.Probe((tag ?? "") + " theme resolve err: " + ex.Message);
                return false;
            }
        }

        private static Office.ThemeColorScheme TryGetThemeColorScheme(PowerPoint.Shape shape, out string source)
        {
            source = null;
            try
            {
                PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
                if (slide != null)
                {
                    try
                    {
                        Office.ThemeColorScheme s = slide.Design.SlideMaster.Theme.ThemeColorScheme;
                        source = "slide.Design.SlideMaster";
                        return s;
                    }
                    catch (Exception ex)
                    {
                        source = "designFail:" + ex.Message;
                    }

                    PowerPoint.Presentation pres = slide.Parent as PowerPoint.Presentation;
                    if (pres != null)
                    {
                        source = "pres.SlideMaster";
                        return pres.SlideMaster.Theme.ThemeColorScheme;
                    }
                }
            }
            catch (Exception ex)
            {
                source = "ex:" + ex.Message;
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
