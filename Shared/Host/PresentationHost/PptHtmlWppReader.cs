using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlWppReader
    {
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
            return TryRead(presentation, slideId, channelId, kind, null, out result, out error);
        }

        public static bool TryRead(
            object presentation,
            string slideId,
            string channelId,
            string kind,
            string shapeId,
            out PptHtmlReadResult result,
            out string error)
        {
            return TryRead(presentation, slideId, channelId, kind, shapeId, false, out result, out error);
        }

        public static bool TryRead(
            object presentation,
            string slideId,
            string channelId,
            string kind,
            string shapeId,
            bool fullPage,
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
            var fontDbg = new PptHtmlReadDebug();
            bool isSkeleton = true;
            fontDbg.Line("begin slide_id=" + trimmed + " host=wpp"
                + (fullPage ? " fullPage=true" : ""));
            try
            {
                object shapesObj = WppCom.GetProperty(slide, "Shapes");
                if (fullPage)
                {
                    if (!CollectShapes(
                        shapesObj,
                        trimmed,
                        slideWidth,
                        slideHeight,
                        0,
                        0,
                        slideWidth,
                        slideHeight,
                        GroupReadMode.FullPage,
                        1,
                        shapes,
                        ref truncated,
                        ref truncatedReason,
                        fontDbg,
                        out string collectFullError))
                    {
                        error = collectFullError;
                        return false;
                    }

                    isSkeleton = false;
                }
                else if (!string.IsNullOrWhiteSpace(shapeId))
                {
                    if (!TryReadFocused(
                        slide,
                        trimmed,
                        shapeId.Trim(),
                        slideWidth,
                        slideHeight,
                        shapes,
                        ref truncated,
                        ref truncatedReason,
                        fontDbg,
                        out string focusError,
                        out isSkeleton))
                    {
                        error = focusError;
                        return false;
                    }
                }
                else if (!CollectShapes(
                    shapesObj,
                    trimmed,
                    slideWidth,
                    slideHeight,
                    0,
                    0,
                    slideWidth,
                    slideHeight,
                    GroupReadMode.Skeleton,
                    1,
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
                ShapeCount = PptHtmlGeom.CountNodes(shapes),
                IsSkeleton = isSkeleton,
                DepthCappedShapeIds = isSkeleton
                    ? PptConventionHtml.CollectDepthCappedIds(shapes)
                    : null,
                DebugFilename = debugFile,
                DebugTrace = new List<string>(fontDbg.Lines)
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

        private enum GroupReadMode
        {
            ShellOnly,
            Skeleton,
            FullTree,
            FullPage
        }

        private static bool CollectShapes(
            object shapes,
            string slideId,
            double slideWidth,
            double slideHeight,
            double parentLeft,
            double parentTop,
            double parentWidth,
            double parentHeight,
            GroupReadMode mode,
            int expandLayer,
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
                count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
            }
            catch (Exception)
            {
                return true;
            }

            for (int i = 1; i <= count; i++)
            {
                if (mode != GroupReadMode.FullPage
                    && PptHtmlGeom.CountNodes(output) >= PptHtmlReadResult.MaxShapes)
                {
                    truncated = true;
                    truncatedReason = "单页形状超过 " + PptHtmlReadResult.MaxShapes + "，已截断";
                    return true;
                }

                object shape = WppCom.GetIndexed(shapes, i);
                if (shape == null)
                {
                    continue;
                }

                if (!AppendNode(
                    shape,
                    slideId,
                    slideWidth,
                    slideHeight,
                    parentLeft,
                    parentTop,
                    parentWidth,
                    parentHeight,
                    mode,
                    expandLayer,
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

        private static bool TryReadFocused(
            object slide,
            string slideId,
            string shapeId,
            double slideWidth,
            double slideHeight,
            List<PptHtmlShapeNode> output,
            ref bool truncated,
            ref string truncatedReason,
            PptHtmlReadDebug fontDbg,
            out string error,
            out bool isSkeleton)
        {
            error = null;
            isSkeleton = false;
            string sidIn = null;
            int comId;
            if (PptShapeId.TryParseShape(shapeId, out sidIn, out comId))
            {
            }
            else if (!PptShapeId.TryParseShapeComId(shapeId, out comId))
            {
                error = "非法 ShapeId: " + shapeId;
                return false;
            }

            if (!string.IsNullOrEmpty(sidIn) && !string.Equals(sidIn, slideId, StringComparison.Ordinal))
            {
                error = "shape_id 与 slide_id 不在同一页";
                return false;
            }

            var path = new List<object>();
            object shapesObj = WppCom.GetProperty(slide, "Shapes");
            if (!TryFindShapePath(shapesObj, comId, path))
            {
                error = "页上找不到 ShapeId: " + shapeId;
                return false;
            }

            object target = path[path.Count - 1];
            string typeName = PeekTypeName(target);
            isSkeleton = typeName == "group";
            GroupReadMode mode = isSkeleton ? GroupReadMode.Skeleton : GroupReadMode.ShellOnly;
            var built = new List<PptHtmlShapeNode>();
            double pL = 0;
            double pT = 0;
            double pW = slideWidth;
            double pH = slideHeight;
            if (path.Count >= 2)
            {
                TryReadBox(path[path.Count - 2], out pL, out pT, out pW, out pH);
            }

            if (!AppendNode(
                target,
                slideId,
                slideWidth,
                slideHeight,
                pL,
                pT,
                pW,
                pH,
                mode,
                0,
                built,
                ref truncated,
                ref truncatedReason,
                fontDbg,
                out error))
            {
                return false;
            }

            if (built.Count == 0)
            {
                error = "无法读取 ShapeId: " + shapeId;
                return false;
            }

            PptHtmlShapeNode current = built[0];
            for (int i = path.Count - 2; i >= 0; i--)
            {
                double aL = 0;
                double aT = 0;
                double aW = slideWidth;
                double aH = slideHeight;
                if (i >= 1)
                {
                    TryReadBox(path[i - 1], out aL, out aT, out aW, out aH);
                }

                if (!AppendNode(
                    path[i],
                    slideId,
                    slideWidth,
                    slideHeight,
                    aL,
                    aT,
                    aW,
                    aH,
                    GroupReadMode.ShellOnly,
                    0,
                    new List<PptHtmlShapeNode>(),
                    ref truncated,
                    ref truncatedReason,
                    fontDbg,
                    out error,
                    out PptHtmlShapeNode ancestor))
                {
                    return false;
                }

                if (ancestor == null)
                {
                    error = "无法读取祖先组";
                    return false;
                }

                ancestor.Children = new List<PptHtmlShapeNode> { current };
                current = ancestor;
            }

            output.Add(current);
            return true;
        }

        private static bool TryFindShapePath(object shapes, int id, List<object> path)
        {
            if (shapes == null)
            {
                return false;
            }

            int count;
            try
            {
                count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
            }
            catch (Exception)
            {
                return false;
            }

            for (int i = 1; i <= count; i++)
            {
                object shape = WppCom.GetIndexed(shapes, i);
                if (shape != null && TryFindShapePathFrom(shape, id, path))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryFindShapePathFrom(object shape, int id, List<object> path)
        {
            try
            {
                path.Add(shape);
                if (Convert.ToInt32(WppCom.GetProperty(shape, "Id")) == id)
                {
                    return true;
                }

                if (TryGetInt(shape, "Type") == 6)
                {
                    object items = WppCom.GetProperty(shape, "GroupItems");
                    int n = Convert.ToInt32(WppCom.GetProperty(items, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object child = WppCom.GetIndexed(items, i);
                        if (child != null && TryFindShapePathFrom(child, id, path))
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            if (path.Count > 0)
            {
                path.RemoveAt(path.Count - 1);
            }

            return false;
        }

        private static string PeekTypeName(object shape)
        {
            int shapeType = TryGetInt(shape, "Type");
            int? autoType = TryGetNullableInt(shape, "AutoShapeType");
            return PptShapeTypeMap.FromShapeType(shapeType, autoType, null);
        }

        private static void TryReadBox(
            object shape,
            out double left,
            out double top,
            out double width,
            out double height)
        {
            left = top = 0;
            width = height = 1;
            try
            {
                left = Convert.ToDouble(WppCom.GetProperty(shape, "Left"));
                top = Convert.ToDouble(WppCom.GetProperty(shape, "Top"));
                width = Convert.ToDouble(WppCom.GetProperty(shape, "Width"));
                height = Convert.ToDouble(WppCom.GetProperty(shape, "Height"));
            }
            catch (Exception)
            {
            }
        }

        private static bool TryReadGroupChildren(
            object group,
            string slideId,
            double slideWidth,
            double slideHeight,
            GroupReadMode childMode,
            int childExpandLayer,
            List<PptHtmlShapeNode> output,
            ref bool pageTextTruncated,
            ref string truncatedReason,
            PptHtmlReadDebug fontDbg,
            out string error)
        {
            error = null;
            object items;
            int count;
            try
            {
                items = WppCom.GetProperty(group, "GroupItems");
                count = Convert.ToInt32(WppCom.GetProperty(items, "Count"));
            }
            catch (Exception)
            {
                return false;
            }

            if (items == null || count <= 0)
            {
                return false;
            }

            TryReadBox(group, out double gL, out double gT, out double gW, out double gH);
            for (int i = 1; i <= count; i++)
            {
                object child = WppCom.GetIndexed(items, i);
                if (child == null)
                {
                    continue;
                }

                if (!AppendNode(
                    child,
                    slideId,
                    slideWidth,
                    slideHeight,
                    gL,
                    gT,
                    gW,
                    gH,
                    childMode,
                    childExpandLayer,
                    output,
                    ref pageTextTruncated,
                    ref truncatedReason,
                    fontDbg,
                    out error))
                {
                    return error == null;
                }
            }

            return true;
        }

        private static bool AppendNode(
            object shape,
            string slideId,
            double slideWidth,
            double slideHeight,
            double parentLeft,
            double parentTop,
            double parentWidth,
            double parentHeight,
            GroupReadMode mode,
            int expandLayer,
            List<PptHtmlShapeNode> output,
            ref bool pageTextTruncated,
            ref string truncatedReason,
            PptHtmlReadDebug fontDbg,
            out string error)
        {
            return AppendNode(
                shape,
                slideId,
                slideWidth,
                slideHeight,
                parentLeft,
                parentTop,
                parentWidth,
                parentHeight,
                mode,
                expandLayer,
                output,
                ref pageTextTruncated,
                ref truncatedReason,
                fontDbg,
                out error,
                out _);
        }

        private static bool AppendNode(
            object shape,
            string slideId,
            double slideWidth,
            double slideHeight,
            double parentLeft,
            double parentTop,
            double parentWidth,
            double parentHeight,
            GroupReadMode mode,
            int expandLayer,
            List<PptHtmlShapeNode> output,
            ref bool pageTextTruncated,
            ref string truncatedReason,
            PptHtmlReadDebug fontDbg,
            out string error,
            out PptHtmlShapeNode built)
        {
            error = null;
            built = null;
            if (output != null
                && mode != GroupReadMode.FullPage
                && PptHtmlGeom.CountNodes(output) >= PptHtmlReadResult.MaxShapes)
            {
                return true;
            }

            int comId;
            try
            {
                comId = Convert.ToInt32(WppCom.GetProperty(shape, "Id"));
            }
            catch (Exception)
            {
                return true;
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
            else if (PptHtmlChartIo.LooksLikeChart(shape))
            {
                typeName = "chart";
            }

            List<PptHtmlShapeNode> groupKids = null;
            bool depthCapped = false;
            if (typeName == "group")
            {
                if (mode == GroupReadMode.FullTree
                    || mode == GroupReadMode.FullPage
                    || (mode == GroupReadMode.Skeleton
                        && expandLayer < PptHtmlReadResult.SkeletonMaxDepth))
                {
                    groupKids = new List<PptHtmlShapeNode>();
                    int childLayer = mode == GroupReadMode.Skeleton ? expandLayer + 1 : 0;
                    if (!TryReadGroupChildren(
                        shape,
                        slideId,
                        slideWidth,
                        slideHeight,
                        mode,
                        childLayer,
                        groupKids,
                        ref pageTextTruncated,
                        ref truncatedReason,
                        fontDbg,
                        out error))
                    {
                        if (error != null)
                        {
                            return false;
                        }

                        typeName = "picture";
                        groupKids = null;
                    }
                }
                else if (mode == GroupReadMode.Skeleton)
                {
                    depthCapped = TryPeekGroupHasChildren(shape);
                }
            }

            if (error != null)
            {
                return false;
            }

            string rasterizedFrom = null;
            if (typeName == "picture" && PeekTypeName(shape) == "group")
            {
                rasterizedFrom = "group";
            }
            else if (PptShapeTypeMap.ShouldRasterizeAsPicture(typeName))
            {
                rasterizedFrom = typeName;
                typeName = "picture";
            }

            bool slim = mode == GroupReadMode.Skeleton;
            string text = "";
            bool textTruncated = false;
            string innerHtml = null;
            PptHtmlChartFormat chartFormat = null;
            if (slim)
            {
                if (typeName != "picture" && typeName != "group" && !PptShapeTypeMap.IsNonEditable(typeName))
                {
                    text = TryReadText(shape);
                    text = PptConventionHtml.TruncateSkeletonText(text, out textTruncated);
                }
            }
            else if (typeName == "chart")
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

            if (textTruncated && !slim)
            {
                pageTextTruncated = true;
            }

            string style = TryBuildStyle(shape, parentLeft, parentTop, parentWidth, parentHeight);
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

            bool editable = slim || !PptShapeTypeMap.IsNonEditable(typeName);
            string tag = PptShapeTypeMap.PreferTag(typeName, !string.IsNullOrEmpty(text));
            string fill = null;
            string fontColor = null;
            double? fontSize = null;
            bool? fontBold = null;
            string fontName = null;
            string lineColor = null;
            double? lineWidth = null;
            string align = null;
            string valign = null;
            double? textWidth = null;
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
            if (!slim && typeName != "picture" && typeName != "media" && typeName != "chart")
            {
                fill = TryReadFill(shape);
                if (typeName != "table")
                {
                    fontColor = TryReadFontColor(shape, fontDbg);
                    fontSize = TryReadFontSize(shape);
                    fontBold = TryReadFontBold(shape);
                    fontName = TryReadFontName(shape);
                    PptHtmlParagraphIo.Snapshot para = PptHtmlParagraphIo.TryReadFromWppShape(shape);
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

                    valign = PptHtmlValignIo.TryReadWpp(shape);
                    if (PptHtmlTextWidthIo.TryMeasureWpp(shape, (float)slideWidth, out double tw))
                    {
                        textWidth = tw;
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

            built = new PptHtmlShapeNode
            {
                ShapeId = "sid" + slideId + "-s" + comId.ToString(CultureInfo.InvariantCulture),
                ShapeType = typeName,
                Tag = typeName == "group" ? "div" : tag,
                Style = style,
                Text = typeName == "group" ? "" : text,
                InnerHtml = innerHtml,
                Editable = editable,
                Fill = typeName == "group" ? null : fill,
                FontColor = fontColor,
                FontSizePt = fontSize,
                FontBold = fontBold,
                FontName = fontName,
                Z = z,
                LineColor = typeName == "group" ? null : lineColor,
                LineWidthPt = typeName == "group" ? null : lineWidth,
                Align = align,
                Valign = valign,
                TextWidthPct = textWidth,
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
                Name = slim ? null : (typeName == "picture" ? name : null),
                Rotation = rotation,
                TextTruncated = textTruncated,
                DepthCapped = depthCapped,
                ChartFormat = chartFormat,
                Children = groupKids
            };
            if (output != null)
            {
                output.Add(built);
            }

            return true;
        }

        private static bool TryPeekGroupHasChildren(object group)
        {
            try
            {
                object items = WppCom.GetProperty(group, "GroupItems");
                int count = Convert.ToInt32(WppCom.GetProperty(items, "Count"));
                return count > 0;
            }
            catch (Exception)
            {
                return true;
            }
        }

        private static void TryReadTextMargins(
            object shape,
            out double? left,
            out double? right,
            out double? top,
            out double? bottom)
        {
            left = right = top = bottom = null;
            try
            {
                if (!IsTruthy(WppCom.GetProperty(shape, "HasTextFrame")))
                {
                    return;
                }

                object tf = WppCom.GetProperty(shape, "TextFrame");
                if (tf == null)
                {
                    return;
                }

                object v;
                v = WppCom.GetProperty(tf, "MarginLeft");
                if (v != null)
                {
                    left = Convert.ToDouble(v);
                }

                v = WppCom.GetProperty(tf, "MarginRight");
                if (v != null)
                {
                    right = Convert.ToDouble(v);
                }

                v = WppCom.GetProperty(tf, "MarginTop");
                if (v != null)
                {
                    top = Convert.ToDouble(v);
                }

                v = WppCom.GetProperty(tf, "MarginBottom");
                if (v != null)
                {
                    bottom = Convert.ToDouble(v);
                }
            }
            catch (Exception)
            {
            }
        }

        private static double? TryReadFontSize(object shape)
        {
            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                object size = font == null ? null : WppCom.GetProperty(font, "Size");
                if (size == null)
                {
                    return null;
                }

                double v = Convert.ToDouble(size);
                return v > 0 ? v : (double?)null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool? TryReadFontBold(object shape)
        {
            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                object bold = font == null ? null : WppCom.GetProperty(font, "Bold");
                if (bold == null)
                {
                    return null;
                }

                int v = Convert.ToInt32(bold);
                if (v == -2)
                {
                    // msoTriStateMixed
                    return null;
                }

                return IsTruthy(bold);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadFontName(object shape)
        {
            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                if (font == null)
                {
                    return null;
                }

                string name = TryGetFontNameProp(font, "NameFarEast");
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = TryGetFontNameProp(font, "Name");
                }

                if (string.IsNullOrWhiteSpace(name))
                {
                    name = TryGetFontNameProp(font, "NameAscii");
                }

                return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryGetFontNameProp(object font, string prop)
        {
            try
            {
                object v = WppCom.GetProperty(font, prop);
                return v == null ? null : Convert.ToString(v);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int? TryReadZ(object shape)
        {
            try
            {
                object z = WppCom.GetProperty(shape, "ZOrderPosition");
                if (z == null)
                {
                    return null;
                }

                int v = Convert.ToInt32(z);
                return v < 0 ? (int?)null : v;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadLineColor(object shape)
        {
            try
            {
                object line = WppCom.GetProperty(shape, "Line");
                if (line == null)
                {
                    return null;
                }

                object visible = WppCom.GetProperty(line, "Visible");
                if (visible != null && !IsTruthy(visible))
                {
                    return "none";
                }

                object fore = WppCom.GetProperty(line, "ForeColor");
                object rgbObj = fore == null ? null : WppCom.GetProperty(fore, "RGB");
                if (rgbObj == null)
                {
                    return null;
                }

                return PptHtmlStyleIo.FormatOfficeRgb(Convert.ToInt32(rgbObj));
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static double? TryReadLineWidth(object shape)
        {
            try
            {
                object line = WppCom.GetProperty(shape, "Line");
                if (line == null)
                {
                    return null;
                }

                object visible = WppCom.GetProperty(line, "Visible");
                if (visible != null && !IsTruthy(visible))
                {
                    return null;
                }

                object w = WppCom.GetProperty(line, "Weight");
                if (w == null)
                {
                    return null;
                }

                double v = Convert.ToDouble(w);
                return v < 0 ? (double?)null : v;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadFill(object shape)
        {
            try
            {
                object fill = WppCom.GetProperty(shape, "Fill");
                if (fill == null)
                {
                    return null;
                }

                object visible = WppCom.GetProperty(fill, "Visible");
                if (visible != null && !IsTruthy(visible))
                {
                    return "none";
                }

                object fore = WppCom.GetProperty(fill, "ForeColor");
                object rgbObj = fore == null ? null : WppCom.GetProperty(fore, "RGB");
                if (rgbObj == null)
                {
                    return null;
                }

                return PptHtmlStyleIo.FormatOfficeRgb(Convert.ToInt32(rgbObj));
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadFontColor(object shape, PptHtmlReadDebug dbg)
        {
            string tag = WppShapeProbeTag(shape);
            try
            {
                try
                {
                    object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                    object tr2 = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
                    object font2 = tr2 == null ? null : WppCom.GetProperty(tr2, "Font");
                    object fill2 = font2 == null ? null : WppCom.GetProperty(font2, "Fill");
                    object fore2 = fill2 == null ? null : WppCom.GetProperty(fill2, "ForeColor");
                    if (fore2 != null)
                    {
                        object rgb2 = WppCom.GetProperty(fore2, "RGB");
                        object type2Obj = WppCom.GetProperty(fore2, "Type");
                        object themeObj = WppCom.GetProperty(fore2, "ObjectThemeColor");
                        object brightObj = WppCom.GetProperty(fore2, "Brightness");
                        int rgbVal = rgb2 == null ? 0 : Convert.ToInt32(rgb2);
                        int rgbNorm = rgbVal & 0x00FFFFFF;
                        int typeVal = type2Obj == null ? 1 : Convert.ToInt32(type2Obj);
                        int themeIdx = themeObj == null ? 0 : Convert.ToInt32(themeObj);
                        float brightness = brightObj == null ? 0f : Convert.ToSingle(brightObj);

                        dbg?.Probe(
                            tag
                            + " tf2-raw type=" + typeVal.ToString(CultureInfo.InvariantCulture)
                            + " rgbRaw=" + rgbVal.ToString(CultureInfo.InvariantCulture)
                            + " rgbNorm=" + rgbNorm.ToString(CultureInfo.InvariantCulture)
                            + " theme=" + themeIdx.ToString(CultureInfo.InvariantCulture)
                            + " bright=" + brightness.ToString("0.###", CultureInfo.InvariantCulture));

                        if (IsMixedFontColorSignal(typeVal, themeIdx, rgbVal))
                        {
                            string mixedHex = TryReadWppFontColorFromFirstCharacter(shape, dbg, tag);
                            if (!string.IsNullOrEmpty(mixedHex))
                            {
                                return mixedHex;
                            }

                            dbg?.Probe(tag + " mixed-char miss -> continue");
                        }

                        bool hasTheme = themeIdx > 0;

                        if (rgbNorm != 0 && !hasTheme)
                        {
                            string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                            dbg?.Probe(tag + " RESULT via=tf2 hex=" + hex);
                            return hex;
                        }

                        if (hasTheme || rgbNorm == 0)
                        {
                            if (hasTheme
                                && TryResolveWppThemeFontRgb(shape, fore2, themeIdx, brightness, out int themeRgb)
                                && (themeRgb & 0x00FFFFFF) != 0)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(themeRgb);
                                dbg?.Probe(tag + " RESULT via=theme hex=" + hex);
                                return hex;
                            }

                            if (hasTheme)
                            {
                                dbg?.Probe(
                                    tag
                                    + " theme-resolve FAIL/zero theme="
                                    + themeIdx.ToString(CultureInfo.InvariantCulture));
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

                        if (typeVal == 1)
                        {
                            string hex = PptHtmlStyleIo.FormatOfficeRgb(0);
                            dbg?.Probe(tag + " RESULT via=tf2-black hex=" + hex);
                            return hex;
                        }
                    }
                }
                catch (Exception ex)
                {
                    dbg?.Probe(tag + " tf2 err=" + ex.Message);
                }

                string charHex = TryReadWppFontColorFromFirstCharacter(shape, dbg, tag);
                if (!string.IsNullOrEmpty(charHex))
                {
                    return charHex;
                }

                if (!IsTruthy(WppCom.GetProperty(shape, "HasTextFrame")))
                {
                    dbg?.Probe(tag + " RESULT via=no-textframe hex=null");
                    return null;
                }

                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                object color = font == null ? null : WppCom.GetProperty(font, "Color");
                object rgbObj = color == null ? null : WppCom.GetProperty(color, "RGB");
                object typeObj = color == null ? null : WppCom.GetProperty(color, "Type");
                object schemeObj = color == null ? null : WppCom.GetProperty(color, "SchemeColor");
                if (rgbObj == null)
                {
                    dbg?.Probe(tag + " RESULT via=tf1-null hex=null");
                    return null;
                }

                int rgb = Convert.ToInt32(rgbObj) & 0x00FFFFFF;
                int type = typeObj == null ? 1 : Convert.ToInt32(typeObj);
                int scheme = schemeObj == null ? -1 : Convert.ToInt32(schemeObj);
                int theme1 = 0;
                try
                {
                    object th = WppCom.GetProperty(color, "ObjectThemeColor");
                    if (th != null)
                    {
                        theme1 = Convert.ToInt32(th);
                    }
                }
                catch (Exception)
                {
                }

                dbg?.Probe(
                    tag
                    + " tf1-raw type=" + type.ToString(CultureInfo.InvariantCulture)
                    + " rgbNorm=" + rgb.ToString(CultureInfo.InvariantCulture)
                    + " scheme=" + scheme.ToString(CultureInfo.InvariantCulture)
                    + " theme=" + theme1.ToString(CultureInfo.InvariantCulture));

                if (IsMixedFontColorSignal(type, theme1, Convert.ToInt32(rgbObj)))
                {
                    string mixedHex = TryReadWppFontColorFromFirstCharacter(shape, dbg, tag);
                    if (!string.IsNullOrEmpty(mixedHex))
                    {
                        return mixedHex;
                    }
                }

                if (rgb == 0 && type != 1)
                {
                    dbg?.Probe(tag + " RESULT via=tf1-skip0 hex=null");
                    return null;
                }

                string hex1 = PptHtmlStyleIo.FormatOfficeRgb(rgb);
                dbg?.Probe(tag + " RESULT via=tf1 hex=" + hex1);
                return hex1;
            }
            catch (Exception ex)
            {
                dbg?.Probe(tag + " RESULT via=exception hex=null err=" + ex.Message);
                return null;
            }
        }

        private static bool IsMixedFontColorSignal(int type, int themeIdx, int rgbRaw)
        {
            return type == -2 || themeIdx == -2 || rgbRaw == int.MinValue;
        }

        private static string TryReadWppFontColorFromFirstCharacter(
            object shape,
            PptHtmlReadDebug dbg,
            string tag)
        {
            try
            {
                object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                object tr2 = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
                if (tr2 != null)
                {
                    object lenObj = WppCom.GetProperty(tr2, "Length");
                    int len = lenObj == null ? 0 : Convert.ToInt32(lenObj);
                    if (len >= 1)
                    {
                        object ch = null;
                        try
                        {
                            ch = WppCom.Invoke(tr2, "Characters", 1, 1);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                ch = WppCom.GetIndexed(tr2, 1);
                            }
                            catch (Exception)
                            {
                            }
                        }

                        object font = ch == null ? null : WppCom.GetProperty(ch, "Font");
                        object fill = font == null ? null : WppCom.GetProperty(font, "Fill");
                        object fore = fill == null ? null : WppCom.GetProperty(fill, "ForeColor");
                        if (fore != null)
                        {
                            object rgb2 = WppCom.GetProperty(fore, "RGB");
                            object type2Obj = WppCom.GetProperty(fore, "Type");
                            object themeObj = WppCom.GetProperty(fore, "ObjectThemeColor");
                            object brightObj = WppCom.GetProperty(fore, "Brightness");
                            int rgbVal = rgb2 == null ? 0 : Convert.ToInt32(rgb2);
                            int rgbNorm = rgbVal & 0x00FFFFFF;
                            int typeVal = type2Obj == null ? 1 : Convert.ToInt32(type2Obj);
                            int themeIdx = themeObj == null ? 0 : Convert.ToInt32(themeObj);
                            float brightness = brightObj == null ? 0f : Convert.ToSingle(brightObj);

                            dbg?.Probe(
                                tag
                                + " char1-tf2 type=" + typeVal.ToString(CultureInfo.InvariantCulture)
                                + " rgbNorm=" + rgbNorm.ToString(CultureInfo.InvariantCulture)
                                + " theme=" + themeIdx.ToString(CultureInfo.InvariantCulture)
                                + " bright=" + brightness.ToString("0.###", CultureInfo.InvariantCulture));

                            if (!IsMixedFontColorSignal(typeVal, themeIdx, rgbVal))
                            {
                                if (rgbNorm != 0 && themeIdx <= 0)
                                {
                                    string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                                    dbg?.Probe(tag + " RESULT via=char1-tf2 hex=" + hex);
                                    return hex;
                                }

                                if (themeIdx > 0
                                    && TryResolveWppThemeFontRgb(shape, fore, themeIdx, brightness, out int themeRgb)
                                    && (themeRgb & 0x00FFFFFF) != 0)
                                {
                                    string hex = PptHtmlStyleIo.FormatOfficeRgb(themeRgb);
                                    dbg?.Probe(tag + " RESULT via=char1-theme hex=" + hex);
                                    return hex;
                                }

                                if (rgbNorm != 0)
                                {
                                    string hex = PptHtmlStyleIo.FormatOfficeRgb(rgbNorm);
                                    dbg?.Probe(tag + " RESULT via=char1-tf2-fallback hex=" + hex);
                                    return hex;
                                }
                            }
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
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                if (tr != null)
                {
                    object lenObj = WppCom.GetProperty(tr, "Length");
                    int len = lenObj == null ? 0 : Convert.ToInt32(lenObj);
                    if (len >= 1)
                    {
                        object ch = WppCom.Invoke(tr, "Characters", 1, 1);
                        object font = ch == null ? null : WppCom.GetProperty(ch, "Font");
                        object color = font == null ? null : WppCom.GetProperty(font, "Color");
                        object rgbObj = color == null ? null : WppCom.GetProperty(color, "RGB");
                        if (rgbObj != null)
                        {
                            int rgb = Convert.ToInt32(rgbObj) & 0x00FFFFFF;
                            int type = 1;
                            try
                            {
                                object t = WppCom.GetProperty(color, "Type");
                                if (t != null)
                                {
                                    type = Convert.ToInt32(t);
                                }
                            }
                            catch (Exception)
                            {
                            }

                            dbg?.Probe(
                                tag
                                + " char1-tf1 type=" + type.ToString(CultureInfo.InvariantCulture)
                                + " rgbNorm=" + rgb.ToString(CultureInfo.InvariantCulture));
                            if (rgb != 0 || type == 1)
                            {
                                string hex = PptHtmlStyleIo.FormatOfficeRgb(rgb);
                                dbg?.Probe(tag + " RESULT via=char1-tf1 hex=" + hex);
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

        private static string WppShapeProbeTag(object shape)
        {
            string id = "?";
            string name = "";
            string text = "";
            try
            {
                id = Convert.ToString(WppCom.GetProperty(shape, "Id")) ?? "?";
            }
            catch (Exception)
            {
            }

            try
            {
                name = Convert.ToString(WppCom.GetProperty(shape, "Name")) ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                text = tr == null ? "" : (Convert.ToString(WppCom.GetProperty(tr, "Text")) ?? "");
                text = text.Replace("\r", " ").Replace("\n", " ").Trim();
                if (text.Length > 20)
                {
                    text = text.Substring(0, 20);
                }
            }
            catch (Exception)
            {
            }

            return "font id=" + id + " name=" + name + " text=[" + text + "]";
        }

        private static bool TryResolveWppThemeFontRgb(
            object shape,
            object fore,
            int themeIdx,
            float brightness,
            out int rgb)
        {
            rgb = 0;
            try
            {
                if (themeIdx == 0)
                {
                    themeIdx = Convert.ToInt32(WppCom.GetProperty(fore, "ObjectThemeColor") ?? 0);
                }

                if (themeIdx == 0
                    || !PptHtmlStyleIo.TryMapThemeColorIndexToSchemeIndex(themeIdx, out int schemeIdx))
                {
                    return false;
                }

                object scheme = TryGetWppThemeColorScheme(shape);
                if (scheme == null)
                {
                    return false;
                }

                object themeColor = WppCom.Invoke(scheme, "Colors", schemeIdx);
                object rgbObj = themeColor == null ? null : WppCom.GetProperty(themeColor, "RGB");
                if (rgbObj == null)
                {
                    return false;
                }

                rgb = Convert.ToInt32(rgbObj) & 0x00FFFFFF;
                if (Math.Abs(brightness) > 0.0001f)
                {
                    int r = rgb & 0xFF;
                    int g = (rgb >> 8) & 0xFF;
                    int b = (rgb >> 16) & 0xFF;
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
                    rgb = r | (g << 8) | (b << 16);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static object TryGetWppThemeColorScheme(object shape)
        {
            try
            {
                object parent = WppCom.GetProperty(shape, "Parent");
                object design = parent == null ? null : WppCom.GetProperty(parent, "Design");
                object master = design == null ? null : WppCom.GetProperty(design, "SlideMaster");
                object theme = master == null ? null : WppCom.GetProperty(master, "Theme");
                object scheme = theme == null ? null : WppCom.GetProperty(theme, "ThemeColorScheme");
                if (scheme != null)
                {
                    return scheme;
                }

                object pres = parent == null ? null : WppCom.GetProperty(parent, "Parent");
                object sm = pres == null ? null : WppCom.GetProperty(pres, "SlideMaster");
                object th = sm == null ? null : WppCom.GetProperty(sm, "Theme");
                return th == null ? null : WppCom.GetProperty(th, "ThemeColorScheme");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryBuildStyle(
            object shape,
            double parentLeft,
            double parentTop,
            double parentWidth,
            double parentHeight)
        {
            try
            {
                double left = Convert.ToDouble(WppCom.GetProperty(shape, "Left"));
                double top = Convert.ToDouble(WppCom.GetProperty(shape, "Top"));
                double width = Convert.ToDouble(WppCom.GetProperty(shape, "Width"));
                double height = Convert.ToDouble(WppCom.GetProperty(shape, "Height"));
                return PptHtmlGeom.StyleFromSlideBox(
                    left,
                    top,
                    width,
                    height,
                    parentLeft,
                    parentTop,
                    parentWidth,
                    parentHeight);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool TryReadExtractTree(
            object presentation,
            string slideId,
            int comId,
            out PptHtmlShapeNode node,
            out string error)
        {
            node = null;
            error = null;
            if (string.IsNullOrWhiteSpace(slideId)
                || !int.TryParse(slideId.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int slideIdInt))
            {
                error = "slide_id 非法";
                return false;
            }

            object slidesCollection = WppCom.GetProperty(presentation, "Slides");
            object slide = FindSlideById(slidesCollection, slideIdInt);
            if (slide == null)
            {
                error = "幻灯片不存在";
                return false;
            }

            var path = new List<object>();
            if (!TryFindShapePath(WppCom.GetProperty(slide, "Shapes"), comId, path))
            {
                error = "页上找不到可抽节点";
                return false;
            }

            object target = path[path.Count - 1];
            object pageSetup = WppCom.GetProperty(presentation, "PageSetup");
            double slideWidth = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideWidth"));
            double slideHeight = Convert.ToDouble(WppCom.GetProperty(pageSetup, "SlideHeight"));
            TryReadBox(target, out double boxL, out double boxT, out double boxW, out double boxH);
            string typeName = PeekTypeName(target);
            var built = new List<PptHtmlShapeNode>();
            bool truncated = false;
            string truncatedReason = null;
            var fontDbg = new PptHtmlReadDebug();
            GroupReadMode mode = typeName == "group" ? GroupReadMode.FullTree : GroupReadMode.ShellOnly;
            if (!AppendNode(
                target,
                slideId.Trim(),
                slideWidth,
                slideHeight,
                boxL,
                boxT,
                boxW,
                boxH,
                mode,
                0,
                built,
                ref truncated,
                ref truncatedReason,
                fontDbg,
                out error))
            {
                return false;
            }

            if (built.Count == 0)
            {
                error = "无法抽出节点";
                return false;
            }

            node = built[0];
            if (typeName != "group")
            {
                node.Style = PptConventionHtml.BuildStyle(0, 0, 100, 100);
            }

            return true;
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
