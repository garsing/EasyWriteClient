using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>WPP 晚绑定 apply；能力弱于 PowerPoint，失败明确报错。</summary>
    internal static class PptHtmlWppApplier
    {
        public static bool TryApply(
            object presentation,
            PptHtmlApplyPlan plan,
            string channelId,
            out PptHtmlApplyResult result,
            out string error)
        {
            result = null;
            error = null;
            if (presentation == null)
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (plan == null || string.IsNullOrEmpty(plan.SlideId))
            {
                error = "无效 apply 规划";
                return false;
            }

            if (!int.TryParse(plan.SlideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int slideIdInt))
            {
                error = "非法 slide_id";
                return false;
            }

            object slides = WppCom.GetProperty(presentation, "Slides");
            object slide = FindSlideById(slides, slideIdInt);
            if (slide == null)
            {
                error = "幻灯片不存在: slide_id=" + plan.SlideId;
                return false;
            }

            double slideWidth;
            double slideHeight;
            try
            {
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

            var warnings = new List<string>();
            if (plan.Warnings != null)
            {
                warnings.AddRange(plan.Warnings);
            }

            int updated = 0;
            int created = 0;
            int skipped = 0;
            var createdShapes = new List<Dictionary<string, object>>();
            var zTargets = new List<KeyValuePair<object, int>>();
            object shapes = WppCom.GetProperty(slide, "Shapes");

            try
            {
                if (!TryPreflight(shapes, plan.Nodes, out error))
                {
                    return false;
                }

                foreach (PptHtmlApplyNode node in plan.Nodes)
                {
                    if (node == null)
                    {
                        continue;
                    }

                    if (node.IsCreate)
                    {
                        if (PptHtmlApplyUpsert.ShouldSkipExplicitCreate(node))
                        {
                            PptHtmlApplyUpsert.AddSkipWarning(node, node.ShapeType, warnings);
                            skipped++;
                            continue;
                        }

                        if (!TryCreate(shapes, slide, node, slideWidth, slideHeight, out string newId, out string createError))
                        {
                            error = createError;
                            return false;
                        }

                        RememberCreatedComId(node, newId);
                        TryCollectZ(shapes, node, zTargets);
                        created++;
                        createdShapes.Add(new Dictionary<string, object>
                        {
                            ["shape_type"] = node.ShapeType ?? "",
                            ["shape_id"] = newId
                        });
                    }
                    else
                    {
                        object existing = null;
                        if (node.ShapeComId.HasValue)
                        {
                            existing = FindShapeById(shapes, node.ShapeComId.Value);
                        }

                        if (existing == null)
                        {
                            PptHtmlMissingShapeAction action = PptHtmlApplyUpsert.PlanMissingShape(
                                node, out string plannedType, out string planError);
                            if (action == PptHtmlMissingShapeAction.Fail)
                            {
                                error = planError;
                                return false;
                            }

                            if (action == PptHtmlMissingShapeAction.Skip)
                            {
                                PptHtmlApplyUpsert.AddSkipWarning(node, plannedType, warnings);
                                skipped++;
                                continue;
                            }

                            PptHtmlApplyUpsert.MutateNodeForCreate(node, plannedType, warnings);
                            if (!TryCreate(shapes, slide, node, slideWidth, slideHeight, out string newId, out string createError))
                            {
                                error = createError;
                                return false;
                            }

                            RememberCreatedComId(node, newId);
                            TryCollectZ(shapes, node, zTargets);
                            created++;
                            createdShapes.Add(new Dictionary<string, object>
                            {
                                ["shape_type"] = node.ShapeType ?? "",
                                ["shape_id"] = newId
                            });
                        }
                        else if (!TryUpdate(
                            shapes,
                            slide,
                            node,
                            slideWidth,
                            slideHeight,
                            warnings,
                            out string updateError))
                        {
                            error = updateError;
                            return false;
                        }
                        else
                        {
                            TryCollectZ(shapes, node, zTargets);
                            updated++;
                        }
                    }
                }

                ApplyZOrder(zTargets);
                RelockAllGeometries(shapes, plan.Nodes, slideWidth, slideHeight);
            }
            catch (Exception ex)
            {
                error = "应用 HTML 失败: " + ex.Message;
                return false;
            }

            int index = 0;
            try
            {
                index = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex"));
            }
            catch (Exception)
            {
            }

            result = new PptHtmlApplyResult
            {
                ChannelId = channelId,
                Kind = "wpp",
                SlideId = plan.SlideId,
                Index = index,
                UpdatedCount = updated,
                CreatedCount = created,
                CreatedShapes = createdShapes,
                Warnings = warnings
            };
            if (skipped > 0)
            {
                warnings.Add("skipped_uncreatable=" + skipped);
            }

            return true;
        }

        private static bool TryPreflight(object shapes, List<PptHtmlApplyNode> nodes, out string error)
        {
            error = null;
            if (nodes == null)
            {
                return true;
            }

            var hard = new List<string>();
            foreach (PptHtmlApplyNode node in nodes)
            {
                if (node == null)
                {
                    continue;
                }

                if (node.IsCreate)
                {
                    if (PptHtmlApplyUpsert.ShouldSkipExplicitCreate(node))
                    {
                        continue;
                    }

                    if (!PptHtmlApplyUpsert.TryValidateExplicitCreate(node, out string ve))
                    {
                        hard.Add(ve);
                    }

                    continue;
                }

                object existing = null;
                if (node.ShapeComId.HasValue)
                {
                    existing = FindShapeById(shapes, node.ShapeComId.Value);
                }

                if (existing != null)
                {
                    continue;
                }

                PptHtmlMissingShapeAction action = PptHtmlApplyUpsert.PlanMissingShape(
                    node, out _, out string planError);
                if (action == PptHtmlMissingShapeAction.Fail)
                {
                    hard.Add(planError);
                }
            }

            if (hard.Count == 0)
            {
                return true;
            }

            error = "应用前预检失败（未写入任何形状）：" + string.Join("；", hard);
            return false;
        }

        private static bool TryUpdate(
            object shapes,
            object slide,
            PptHtmlApplyNode node,
            double slideWidth,
            double slideHeight,
            List<string> warnings,
            out string error)
        {
            error = null;
            if (!node.ShapeComId.HasValue)
            {
                error = "更新节点缺少 ShapeId";
                return false;
            }

            object shape = FindShapeById(shapes, node.ShapeComId.Value);
            if (shape == null)
            {
                error = "ShapeId 找不到: " + (node.ShapeId ?? "");
                return false;
            }

            string existingType = ResolveExistingType(shape);

            if (node.TableCells != null)
            {
                if (!TryWriteTable(shape, node.TableCells, out error))
                {
                    return false;
                }
            }
            else if (node.HasText)
            {
                if (existingType == "chart" || existingType == "smartart")
                {
                    warnings.Add("忽略对 " + existingType + " 的文本修改: " + node.ShapeId);
                }
                else if (existingType != "picture" && existingType != "media")
                {
                    if (!TryWriteText(shape, node.Text ?? "", out error))
                    {
                        return false;
                    }
                }
            }

            if (node.HasGeometry)
            {
                TrySet(shape, "Left", node.LeftPct.GetValueOrDefault() / 100.0 * slideWidth);
                TrySet(shape, "Top", node.TopPct.GetValueOrDefault() / 100.0 * slideHeight);
                TrySet(shape, "Width", node.WidthPct.GetValueOrDefault() / 100.0 * slideWidth);
                TrySet(shape, "Height", node.HeightPct.GetValueOrDefault() / 100.0 * slideHeight);
            }

            if (node.Rotation.HasValue)
            {
                TrySet(shape, "Rotation", node.Rotation.Value);
            }

            if (!TryApplyColors(shape, node, existingType, out error))
            {
                return false;
            }

            if (!TryApplyFont(shape, node, existingType, out error))
            {
                return false;
            }

            if (!TryApplyLine(shape, node, existingType, out error))
            {
                return false;
            }

            LockTextFrameAndGeometry(shape, node, slideWidth, slideHeight);

            if (!string.IsNullOrWhiteSpace(node.DataSrc)
                && (existingType == "picture"
                    || existingType == "media"
                    || PptShapeTypeMap.ShouldRasterizeAsPicture(existingType)))
            {
                if (string.IsNullOrEmpty(node.ResolvedLocalPath) || !File.Exists(node.ResolvedLocalPath))
                {
                    error = "替换文件不存在: " + (node.DataSrc ?? "");
                    return false;
                }

                double left = Convert.ToDouble(WppCom.GetProperty(shape, "Left") ?? 0.0);
                double top = Convert.ToDouble(WppCom.GetProperty(shape, "Top") ?? 0.0);
                double width = Convert.ToDouble(WppCom.GetProperty(shape, "Width") ?? 0.0);
                double height = Convert.ToDouble(WppCom.GetProperty(shape, "Height") ?? 0.0);
                try
                {
                    WppCom.Invoke(shape, "Delete");
                }
                catch (Exception ex)
                {
                    error = "删除旧形状失败: " + ex.Message;
                    return false;
                }

                node.IsCreate = true;
                node.HasGeometry = true;
                node.LeftPct = left / slideWidth * 100;
                node.TopPct = top / slideHeight * 100;
                node.WidthPct = width / slideWidth * 100;
                node.HeightPct = height / slideHeight * 100;
                node.ShapeType = existingType == "media" ? "media" : "picture";
                if (!TryCreate(shapes, slide, node, slideWidth, slideHeight, out string newId, out error))
                {
                    return false;
                }

                RememberCreatedComId(node, newId);
                if (PptShapeTypeMap.ShouldRasterizeAsPicture(existingType))
                {
                    warnings?.Add(
                        "已将 " + existingType + " 替换为 picture: " + (node.ShapeId ?? ""));
                }
            }

            return true;
        }

        private static string ResolveExistingType(object shape)
        {
            try
            {
                int st = Convert.ToInt32(WppCom.GetProperty(shape, "Type") ?? 0);
                int? auto = null;
                try
                {
                    object a = WppCom.GetProperty(shape, "AutoShapeType");
                    if (a != null)
                    {
                        auto = Convert.ToInt32(a);
                    }
                }
                catch (Exception)
                {
                }

                int? ph = null;
                try
                {
                    if (st == 14)
                    {
                        object pf = WppCom.GetProperty(shape, "PlaceholderFormat");
                        ph = Convert.ToInt32(WppCom.GetProperty(pf, "Type"));
                    }
                }
                catch (Exception)
                {
                }

                return PptShapeTypeMap.FromShapeType(st, auto, ph);
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool TryCreate(
            object shapes,
            object slide,
            PptHtmlApplyNode node,
            double slideWidth,
            double slideHeight,
            out string newShapeId,
            out string error)
        {
            newShapeId = null;
            error = null;
            double left = node.LeftPct.GetValueOrDefault() / 100.0 * slideWidth;
            double top = node.TopPct.GetValueOrDefault() / 100.0 * slideHeight;
            double width = node.WidthPct.GetValueOrDefault() / 100.0 * slideWidth;
            double height = node.HeightPct.GetValueOrDefault() / 100.0 * slideHeight;
            if (width <= 0)
            {
                width = 100;
            }

            if (height <= 0)
            {
                height = 50;
            }

            object shape = null;
            string type = node.ShapeType ?? "";
            try
            {
                if (type == "textbox")
                {
                    shape = Invoke(shapes, "AddTextbox", 1, left, top, width, height);
                    if (node.HasText)
                    {
                        TryWriteText(shape, node.Text ?? "", out _);
                    }
                }
                else if (type == "table")
                {
                    int rows = Math.Max(1, node.TableCells == null ? 1 : node.TableCells.Count);
                    int cols = node.TableCells == null || node.TableCells.Count == 0
                        ? 1
                        : Math.Max(1, node.TableCells[0].Count);
                    if (rows > 20 || cols > 20)
                    {
                        error = "新建表格不得超过 20×20";
                        return false;
                    }

                    shape = Invoke(shapes, "AddTable", rows, cols, left, top, width, height);
                    if (node.TableCells != null)
                    {
                        TryWriteTable(shape, node.TableCells, out _);
                    }
                }
                else if (type == "picture")
                {
                    if (string.IsNullOrEmpty(node.ResolvedLocalPath) || !File.Exists(node.ResolvedLocalPath))
                    {
                        error = "图片文件不存在: " + (node.DataSrc ?? "")
                            + " local=" + (node.ResolvedLocalPath ?? "");
                        return false;
                    }

                    string picPath = Path.GetFullPath(node.ResolvedLocalPath);
                    try
                    {
                        shape = Invoke(
                            shapes,
                            "AddPicture",
                            picPath,
                            false,
                            true,
                            left,
                            top,
                            width,
                            height);
                    }
                    catch (Exception ex)
                    {
                        error = "插入图片失败: " + ex.Message + "；path=" + picPath;
                        return false;
                    }
                }
                else if (type == "chart" || type == "media")
                {
                    error = "当前 WPS 演示宿主暂无法稳定创建 " + type + "，请用 powerpoint 渠道或仅更新已有形状";
                    return false;
                }
                else if (PptShapeTypeMap.TryGetAutoShapeType(type, out int autoType))
                {
                    shape = Invoke(shapes, "AddShape", autoType, left, top, width, height);
                    if (node.HasText && !string.IsNullOrEmpty(node.Text))
                    {
                        TryWriteText(shape, node.Text, out _);
                    }
                }
                else
                {
                    error = "无法创建 data-shape-type=" + type;
                    return false;
                }
            }
            catch (Exception ex)
            {
                error = "创建形状失败 (" + type + "): " + ex.Message;
                return false;
            }

            if (shape == null)
            {
                error = "创建形状返回空";
                return false;
            }

            if (!TryApplyColors(shape, node, type, out error))
            {
                return false;
            }

            if (!TryApplyFont(shape, node, type, out error))
            {
                return false;
            }

            if (!TryApplyLine(shape, node, type, out error))
            {
                return false;
            }

            LockTextFrameAndGeometry(shape, node, slideWidth, slideHeight);

            try
            {
                string sid = Convert.ToString(WppCom.GetProperty(slide, "SlideID")) ?? "";
                string id = Convert.ToString(WppCom.GetProperty(shape, "Id")) ?? "";
                newShapeId = "sid" + sid + "-s" + id;
            }
            catch (Exception)
            {
                newShapeId = "";
            }

            return true;
        }

        private static void LockTextFrameAndGeometry(
            object shape,
            PptHtmlApplyNode node,
            double slideWidth,
            double slideHeight)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                // PpAutoSizeNone = 0
                TrySet(tf, "AutoSize", 0);
            }
            catch (Exception)
            {
            }

            if (node != null && node.HasGeometry)
            {
                TrySet(shape, "Left", node.LeftPct.GetValueOrDefault() / 100.0 * slideWidth);
                TrySet(shape, "Top", node.TopPct.GetValueOrDefault() / 100.0 * slideHeight);
                TrySet(shape, "Width", node.WidthPct.GetValueOrDefault() / 100.0 * slideWidth);
                TrySet(shape, "Height", node.HeightPct.GetValueOrDefault() / 100.0 * slideHeight);
            }
        }

        private static bool TryApplyFont(object shape, PptHtmlApplyNode node, string shapeType, out string error)
        {
            error = null;
            if (shape == null || node == null)
            {
                return true;
            }

            if (shapeType == "picture" || shapeType == "media" || shapeType == "table")
            {
                return true;
            }

            if (!node.FontSizePt.HasValue && !node.FontBold.HasValue)
            {
                return true;
            }

            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                if (font == null)
                {
                    return true;
                }

                if (node.FontSizePt.HasValue)
                {
                    TrySet(font, "Size", node.FontSizePt.Value);
                }

                if (node.FontBold.HasValue)
                {
                    TrySet(font, "Bold", node.FontBold.Value ? -1 : 0);
                }
            }
            catch (Exception ex)
            {
                error = "写字体样式失败: " + ex.Message;
                return false;
            }

            return true;
        }

        private static bool TryApplyLine(object shape, PptHtmlApplyNode node, string shapeType, out string error)
        {
            error = null;
            if (shape == null || node == null)
            {
                return true;
            }

            if (shapeType == "picture" || shapeType == "media")
            {
                return true;
            }

            if (string.IsNullOrEmpty(node.LineColor) && !node.LineWidthPt.HasValue)
            {
                return true;
            }

            try
            {
                object line = WppCom.GetProperty(shape, "Line");
                if (line == null)
                {
                    return true;
                }

                if (!string.IsNullOrEmpty(node.LineColor))
                {
                    if (string.Equals(node.LineColor, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        TrySet(line, "Visible", 0);
                    }
                    else
                    {
                        if (!PptHtmlStyleIo.TryParseHexToOfficeRgb(node.LineColor, out int rgb, out error))
                        {
                            return false;
                        }

                        TrySet(line, "Visible", -1);
                        object fore = WppCom.GetProperty(line, "ForeColor");
                        TrySet(fore, "RGB", rgb);
                    }
                }

                if (node.LineWidthPt.HasValue)
                {
                    TrySet(line, "Visible", -1);
                    TrySet(line, "Weight", node.LineWidthPt.Value);
                }
            }
            catch (Exception ex)
            {
                error = "写线条样式失败: " + ex.Message;
                return false;
            }

            return true;
        }

        private static void RememberCreatedComId(PptHtmlApplyNode node, string newShapeId)
        {
            if (node == null || string.IsNullOrEmpty(newShapeId))
            {
                return;
            }

            if (PptShapeId.TryParseShape(newShapeId, out _, out int comId))
            {
                node.ShapeComId = comId;
                node.ShapeId = newShapeId;
            }
        }

        private static void TryCollectZ(
            object shapes,
            PptHtmlApplyNode node,
            List<KeyValuePair<object, int>> targets)
        {
            if (shapes == null || node == null || !node.Z.HasValue || targets == null)
            {
                return;
            }

            if (!node.ShapeComId.HasValue)
            {
                return;
            }

            object shape = FindShapeById(shapes, node.ShapeComId.Value);
            if (shape == null)
            {
                return;
            }

            targets.Add(new KeyValuePair<object, int>(shape, node.Z.Value));
        }

        private static void ApplyZOrder(List<KeyValuePair<object, int>> targets)
        {
            if (targets == null || targets.Count == 0)
            {
                return;
            }

            // msoBringToFront = 0
            targets.Sort((a, b) => a.Value.CompareTo(b.Value));
            foreach (KeyValuePair<object, int> item in targets)
            {
                try
                {
                    WppCom.Invoke(item.Key, "ZOrder", 0);
                }
                catch (Exception)
                {
                }
            }
        }

        private static void RelockAllGeometries(
            object shapes,
            List<PptHtmlApplyNode> nodes,
            double slideWidth,
            double slideHeight)
        {
            if (shapes == null || nodes == null)
            {
                return;
            }

            foreach (PptHtmlApplyNode node in nodes)
            {
                if (node == null || !node.HasGeometry || !node.ShapeComId.HasValue)
                {
                    continue;
                }

                object shape = FindShapeById(shapes, node.ShapeComId.Value);
                if (shape == null)
                {
                    continue;
                }

                LockTextFrameAndGeometry(shape, node, slideWidth, slideHeight);
            }
        }

        private static bool TryApplyColors(object shape, PptHtmlApplyNode node, string shapeType, out string error)
        {
            error = null;
            if (shape == null || node == null)
            {
                return true;
            }

            if (shapeType == "picture" || shapeType == "media")
            {
                return true;
            }

            if (!string.IsNullOrEmpty(node.Fill))
            {
                try
                {
                    object fill = WppCom.GetProperty(shape, "Fill");
                    if (fill == null)
                    {
                        return true;
                    }

                    if (string.Equals(node.Fill, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        TrySet(fill, "Visible", 0);
                    }
                    else
                    {
                        if (!PptHtmlStyleIo.TryParseHexToOfficeRgb(node.Fill, out int rgb, out error))
                        {
                            return false;
                        }

                        TrySet(fill, "Visible", -1);
                        try
                        {
                            WppCom.Invoke(fill, "Solid");
                        }
                        catch (Exception)
                        {
                        }

                        object fore = WppCom.GetProperty(fill, "ForeColor");
                        TrySet(fore, "RGB", rgb);
                    }
                }
                catch (Exception ex)
                {
                    error = "写填充色失败: " + ex.Message;
                    return false;
                }
            }

            if (!string.IsNullOrEmpty(node.FontColor) && shapeType != "table")
            {
                try
                {
                    if (!PptHtmlStyleIo.TryParseHexToOfficeRgb(node.FontColor, out int rgb, out error))
                    {
                        return false;
                    }

                    object tf = WppCom.GetProperty(shape, "TextFrame");
                    object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                    object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                    object color = font == null ? null : WppCom.GetProperty(font, "Color");
                    TrySet(color, "RGB", rgb);
                }
                catch (Exception ex)
                {
                    error = "写字体颜色失败: " + ex.Message;
                    return false;
                }
            }

            return true;
        }

        private static bool TryWriteText(object shape, string text, out string error)
        {
            error = null;
            try
            {
                object has = WppCom.GetProperty(shape, "HasTextFrame");
                if (has != null && Convert.ToInt32(has) != -1 && Convert.ToInt32(has) != 1)
                {
                    return true;
                }

                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                TrySet(tr, "Text", text ?? "");
                return true;
            }
            catch (Exception ex)
            {
                error = "写文本失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryWriteTable(object shape, List<List<string>> cells, out string error)
        {
            error = null;
            try
            {
                object table = WppCom.GetProperty(shape, "Table");
                int rows = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(table, "Rows"), "Count"));
                int cols = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(table, "Columns"), "Count"));
                if (cells.Count > rows || (cells.Count > 0 && cells[0].Count > cols))
                {
                    error = "表格行列多于现有表（首期不自动扩表）";
                    return false;
                }

                for (int r = 0; r < cells.Count; r++)
                {
                    for (int c = 0; c < cells[r].Count && c < cols; c++)
                    {
                        object cell = Invoke(table, "Cell", r + 1, c + 1);
                        object cellShape = WppCom.GetProperty(cell, "Shape");
                        TryWriteText(cellShape, cells[r][c] ?? "", out _);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "写表格失败: " + ex.Message;
                return false;
            }
        }

        private static object FindSlideById(object slides, int slideId)
        {
            try
            {
                object found = Invoke(slides, "FindBySlideID", slideId);
                if (found != null)
                {
                    return found;
                }
            }
            catch (Exception)
            {
            }

            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object slide = WppCom.GetIndexed(slides, i);
                try
                {
                    if (Convert.ToInt32(WppCom.GetProperty(slide, "SlideID")) == slideId)
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

        private static object FindShapeById(object shapes, int id)
        {
            int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object shape = WppCom.GetIndexed(shapes, i);
                try
                {
                    if (Convert.ToInt32(WppCom.GetProperty(shape, "Id")) == id)
                    {
                        return shape;
                    }
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static object Invoke(object target, string method, params object[] args)
        {
            return target.GetType().InvokeMember(
                method,
                BindingFlags.InvokeMethod,
                null,
                target,
                args);
        }

        private static void TrySet(object target, string name, object value)
        {
            try
            {
                target.GetType().InvokeMember(
                    name,
                    BindingFlags.SetProperty,
                    null,
                    target,
                    new[] { value });
            }
            catch (Exception)
            {
            }
        }
    }
}
