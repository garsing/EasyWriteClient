using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlPowerPointApplier
    {
        // xlColumnClustered / xlBarClustered / xlLine / xlPie
        private const int XlColumnClustered = 51;
        private const int XlBarClustered = 57;
        private const int XlLine = 4;
        private const int XlPie = 5;

        public static bool TryApply(
            PowerPoint.Presentation presentation,
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
                error = "幻灯片不存在: slide_id=" + plan.SlideId;
                return false;
            }

            float slideWidth;
            float slideHeight;
            try
            {
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

            var warnings = new List<string>();
            if (plan.Warnings != null)
            {
                warnings.AddRange(plan.Warnings);
            }

            int updated = 0;
            int created = 0;
            int skipped = 0;
            var createdShapes = new List<Dictionary<string, object>>();
            var zTargets = new List<KeyValuePair<PowerPoint.Shape, int>>();
            bool debugOn = EasyWriteDiagnostics.IsEnabled(DebugCategory.PptHtml);
            PptHtmlApplyDebug dbg = debugOn ? new PptHtmlApplyDebug() : null;
            if (dbg != null)
            {
                dbg.Line(
                    "begin slide_id=" + plan.SlideId
                    + " slideSize=" + slideWidth.ToString("0.#", CultureInfo.InvariantCulture)
                    + "x" + slideHeight.ToString("0.#", CultureInfo.InvariantCulture)
                    + " nodes=" + (plan.Nodes == null ? 0 : plan.Nodes.Count)
                    + " shapesOnSlideBefore=" + CountShapes(slide));
            }

            try
            {
                // 先预检：避免写到一半因 freeform 等失败留下半成品
                if (!TryPreflight(slide, plan.Nodes, warnings, out error))
                {
                    if (dbg != null)
                    {
                        dbg.Line("preflight FAIL: " + error);
                        result = new PptHtmlApplyResult
                        {
                            ChannelId = channelId,
                            Kind = "ppt",
                            SlideId = plan.SlideId,
                            Warnings = warnings
                        };
                        AttachDebug(result, dbg, plan.SlideId);
                    }

                    return false;
                }

                dbg?.Line("preflight ok");

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
                            dbg?.Step("SKIP_CREATE", node);
                            skipped++;
                            continue;
                        }

                        if (!TryCreate(slide, node, slideWidth, slideHeight, warnings, out string newId, out string createError))
                        {
                            error = createError;
                            if (dbg != null)
                            {
                                dbg.Step("CREATE_FAIL", node, createError);
                                result = FailResult(channelId, plan, warnings);
                                AttachDebug(result, dbg, plan.SlideId);
                            }

                            return false;
                        }

                        RememberCreatedComId(node, newId);
                        TryCollectZ(slide, node, zTargets);
                        dbg?.Step("CREATE", node, "newId=" + newId
                            + (dbg != null ? " " + DescribeShape(slide, node, slideWidth, slideHeight) : ""));
                        created++;
                        createdShapes.Add(new Dictionary<string, object>
                        {
                            ["shape_type"] = node.ShapeType ?? "",
                            ["shape_id"] = newId
                        });
                    }
                    else
                    {
                        PowerPoint.Shape existing = null;
                        if (node.ShapeComId.HasValue)
                        {
                            existing = FindShapeById(slide.Shapes, node.ShapeComId.Value);
                        }

                        if (existing == null)
                        {
                            PptHtmlMissingShapeAction action = PptHtmlApplyUpsert.PlanMissingShape(
                                node, out string plannedType, out string planError);
                            if (action == PptHtmlMissingShapeAction.Fail)
                            {
                                error = planError;
                                if (dbg != null)
                                {
                                    dbg.Step("UPSERT_FAIL", node, planError);
                                    result = FailResult(channelId, plan, warnings);
                                    AttachDebug(result, dbg, plan.SlideId);
                                }

                                return false;
                            }

                            if (action == PptHtmlMissingShapeAction.Skip)
                            {
                                PptHtmlApplyUpsert.AddSkipWarning(node, plannedType, warnings);
                                dbg?.Step("UPSERT_SKIP", node, "planned=" + plannedType);
                                skipped++;
                                continue;
                            }

                            PptHtmlApplyUpsert.MutateNodeForCreate(node, plannedType, warnings);
                            if (!TryCreate(slide, node, slideWidth, slideHeight, warnings, out string newId, out string createError))
                            {
                                error = createError;
                                if (dbg != null)
                                {
                                    dbg.Step("UPSERT_CREATE_FAIL", node, createError);
                                    result = FailResult(channelId, plan, warnings);
                                    AttachDebug(result, dbg, plan.SlideId);
                                }

                                return false;
                            }

                            RememberCreatedComId(node, newId);
                            TryCollectZ(slide, node, zTargets);
                            dbg?.Step(
                                "UPSERT_CREATE",
                                node,
                                "newId=" + newId
                                + (dbg != null ? " " + DescribeShape(slide, node, slideWidth, slideHeight) : ""));
                            created++;
                            createdShapes.Add(new Dictionary<string, object>
                            {
                                ["shape_type"] = node.ShapeType ?? "",
                                ["shape_id"] = newId
                            });
                        }
                        else if (!TryUpdate(
                                slide,
                                node,
                                slideWidth,
                                slideHeight,
                                warnings,
                                out string updateError))
                        {
                            error = updateError;
                            if (dbg != null)
                            {
                                dbg.Step("UPDATE_FAIL", node, updateError);
                                result = FailResult(channelId, plan, warnings);
                                AttachDebug(result, dbg, plan.SlideId);
                            }

                            return false;
                        }
                        else
                        {
                            TryCollectZ(slide, node, zTargets);
                            dbg?.Step(
                                "UPDATE",
                                node,
                                dbg != null ? DescribeShape(slide, node, slideWidth, slideHeight) : null);
                            updated++;
                        }
                    }
                }

                if (dbg != null)
                {
                    dbg.Line("before_zorder zTargets=" + zTargets.Count + " shapesOnSlide=" + CountShapes(slide));
                }

                ApplyZOrder(zTargets);
                dbg?.Line("after_zorder");

                // ZOrder 后再次钉死几何，防止个别 AutoShape 在叠放调整后位置漂移
                RelockAllGeometries(slide, plan.Nodes, slideWidth, slideHeight);
                dbg?.Line("after_relock");

                if (dbg != null)
                {
                    SnapshotSlide(slide, slideWidth, slideHeight, dbg);
                    VerifyNodesStillPresent(slide, plan.Nodes, slideWidth, slideHeight, dbg);
                }
            }
            catch (Exception ex)
            {
                error = "应用 HTML 失败: " + ex.Message;
                if (dbg != null)
                {
                    dbg.Line("EXCEPTION: " + ex);
                    result = new PptHtmlApplyResult
                    {
                        ChannelId = channelId,
                        Kind = "ppt",
                        SlideId = plan.SlideId,
                        Warnings = warnings
                    };
                    AttachDebug(result, dbg, plan.SlideId);
                }

                return false;
            }

            result = new PptHtmlApplyResult
            {
                ChannelId = channelId,
                Kind = "ppt",
                SlideId = plan.SlideId,
                Index = TryGetIndex(slide),
                UpdatedCount = updated,
                CreatedCount = created,
                CreatedShapes = createdShapes,
                Warnings = warnings
            };
            if (skipped > 0)
            {
                warnings.Add("skipped_uncreatable=" + skipped);
            }

            AttachDebug(result, dbg, plan.SlideId);
            return true;
        }

        private static PptHtmlApplyResult FailResult(
            string channelId,
            PptHtmlApplyPlan plan,
            List<string> warnings)
        {
            return new PptHtmlApplyResult
            {
                ChannelId = channelId,
                Kind = "ppt",
                SlideId = plan?.SlideId,
                Warnings = warnings
            };
        }

        private static void AttachDebug(PptHtmlApplyResult result, PptHtmlApplyDebug dbg, string slideId)
        {
            if (dbg == null)
            {
                return;
            }

            string file = dbg.TryWriteToSession(slideId, out string writeErr);
            if (!string.IsNullOrEmpty(writeErr))
            {
                dbg.Line("write_debug_file_fail: " + writeErr);
            }
            else
            {
                dbg.Line("debug_file=" + (file ?? ""));
            }

            // 镜像到统一诊断通道（config.json DebugCategories 含 ppt_html / all 时可见）
            foreach (string line in dbg.Lines)
            {
                EasyWriteDiagnostics.Log(DebugCategory.PptHtml, line);
            }

            if (result != null)
            {
                result.DebugTrace = new List<string>(dbg.Lines);
                result.DebugFilename = file;
            }
        }

        private static int CountShapes(PowerPoint.Slide slide)
        {
            try
            {
                return slide.Shapes.Count;
            }
            catch (Exception)
            {
                return -1;
            }
        }

        private static string DescribeShape(
            PowerPoint.Slide slide,
            PptHtmlApplyNode node,
            float slideWidth,
            float slideHeight)
        {
            if (slide == null || node == null || !node.ShapeComId.HasValue)
            {
                return "shape=?";
            }

            PowerPoint.Shape shape = FindShapeById(slide.Shapes, node.ShapeComId.Value);
            if (shape == null)
            {
                return "shape=MISSING_AFTER_WRITE";
            }

            return DescribeShapeGeom(shape, slideWidth, slideHeight);
        }

        private static string DescribeShapeGeom(PowerPoint.Shape shape, float slideWidth, float slideHeight)
        {
            if (shape == null)
            {
                return "shape=null";
            }

            try
            {
                string fill = "?";
                try
                {
                    fill = shape.Fill.Visible == Office.MsoTriState.msoFalse
                        ? "none"
                        : "vis";
                }
                catch (Exception)
                {
                }

                int z = -1;
                try
                {
                    z = shape.ZOrderPosition;
                }
                catch (Exception)
                {
                }

                return "live=["
                    + Pct(shape.Left, slideWidth) + ","
                    + Pct(shape.Top, slideHeight) + ","
                    + Pct(shape.Width, slideWidth) + ","
                    + Pct(shape.Height, slideHeight)
                    + "] z=" + z.ToString(CultureInfo.InvariantCulture)
                    + " fill=" + fill
                    + " name=" + (shape.Name ?? "");
            }
            catch (Exception ex)
            {
                return "live=err:" + ex.Message;
            }
        }

        private static string TryReadbackFontColor(PowerPoint.Shape shape)
        {
            try
            {
                int rgb = shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB;
                return PptHtmlStyleIo.FormatOfficeRgb(rgb);
            }
            catch (Exception)
            {
            }

            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    return PptHtmlStyleIo.FormatOfficeRgb(shape.TextFrame.TextRange.Font.Color.RGB);
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static string Pct(float value, float total)
        {
            if (total <= 0)
            {
                return "?";
            }

            return (value / total * 100.0).ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static void SnapshotSlide(
            PowerPoint.Slide slide,
            float slideWidth,
            float slideHeight,
            PptHtmlApplyDebug dbg)
        {
            if (slide == null || dbg == null)
            {
                return;
            }

            dbg.Line("--- SNAPSHOT all shapes on slide ---");
            try
            {
                int count = slide.Shapes.Count;
                dbg.Line("shapeCount=" + count.ToString(CultureInfo.InvariantCulture));
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape;
                    try
                    {
                        shape = slide.Shapes[i];
                    }
                    catch (Exception ex)
                    {
                        dbg.Line("  [" + i + "] read_fail: " + ex.Message);
                        continue;
                    }

                    int id = -1;
                    string typeName = "?";
                    try
                    {
                        id = shape.Id;
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        int st = (int)shape.Type;
                        int? auto = null;
                        try
                        {
                            auto = Convert.ToInt32(shape.AutoShapeType);
                        }
                        catch (Exception)
                        {
                        }

                        typeName = PptShapeTypeMap.FromShapeType(st, auto, null);
                    }
                    catch (Exception)
                    {
                    }

                    string band = "";
                    try
                    {
                        double topPct = shape.Top / slideHeight * 100.0;
                        if (topPct < 15)
                        {
                            band = " NAV_BAND";
                        }
                    }
                    catch (Exception)
                    {
                    }

                    dbg.Line(
                        "  #" + i.ToString(CultureInfo.InvariantCulture)
                        + " comId=" + id.ToString(CultureInfo.InvariantCulture)
                        + " type=" + typeName
                        + " " + DescribeShapeGeom(shape, slideWidth, slideHeight)
                        + band);
                }
            }
            catch (Exception ex)
            {
                dbg.Line("snapshot_fail: " + ex.Message);
            }
        }

        private static void VerifyNodesStillPresent(
            PowerPoint.Slide slide,
            List<PptHtmlApplyNode> nodes,
            float slideWidth,
            float slideHeight,
            PptHtmlApplyDebug dbg)
        {
            if (slide == null || nodes == null || dbg == null)
            {
                return;
            }

            dbg.Line("--- VERIFY html nodes still on slide ---");
            foreach (PptHtmlApplyNode node in nodes)
            {
                if (node == null || !node.ShapeComId.HasValue)
                {
                    continue;
                }

                PowerPoint.Shape shape = FindShapeById(slide.Shapes, node.ShapeComId.Value);
                if (shape == null)
                {
                    dbg.Step("MISSING", node, "comId not found after apply");
                    continue;
                }

                string note = DescribeShapeGeom(shape, slideWidth, slideHeight);
                if (!string.IsNullOrEmpty(node.FontColor))
                {
                    note += " html_font=" + node.FontColor
                        + " slide_font=" + (TryReadbackFontColor(shape) ?? "?");
                }

                if (node.HasGeometry)
                {
                    try
                    {
                        double topPct = shape.Top / slideHeight * 100.0;
                        double leftPct = shape.Left / slideWidth * 100.0;
                        double dTop = Math.Abs(topPct - node.TopPct.GetValueOrDefault());
                        double dLeft = Math.Abs(leftPct - node.LeftPct.GetValueOrDefault());
                        if (dTop > 3 || dLeft > 3)
                        {
                            note += " GEO_DRIFT html_top="
                                + node.TopPct.GetValueOrDefault().ToString("0.##", CultureInfo.InvariantCulture)
                                + " html_left="
                                + node.LeftPct.GetValueOrDefault().ToString("0.##", CultureInfo.InvariantCulture);
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                dbg.Step("OK", node, note);
            }
        }

        private static bool TryPreflight(
            PowerPoint.Slide slide,
            List<PptHtmlApplyNode> nodes,
            List<string> warnings,
            out string error)
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

                PowerPoint.Shape existing = null;
                if (node.ShapeComId.HasValue)
                {
                    existing = FindShapeById(slide.Shapes, node.ShapeComId.Value);
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
            PowerPoint.Slide slide,
            PptHtmlApplyNode node,
            float slideWidth,
            float slideHeight,
            List<string> warnings,
            out string error)
        {
            error = null;
            if (!node.ShapeComId.HasValue)
            {
                error = "更新节点缺少 ShapeId";
                return false;
            }

            PowerPoint.Shape shape = FindShapeById(slide.Shapes, node.ShapeComId.Value);
            if (shape == null)
            {
                error = "ShapeId 找不到: " + (node.ShapeId ?? "");
                return false;
            }

            string existingType = "";
            try
            {
                int st = (int)shape.Type;
                int? auto = null;
                try
                {
                    auto = Convert.ToInt32(shape.AutoShapeType);
                }
                catch (Exception)
                {
                }

                int? ph = null;
                try
                {
                    if (st == 14)
                    {
                        ph = Convert.ToInt32(shape.PlaceholderFormat.Type);
                    }
                }
                catch (Exception)
                {
                }

                existingType = PptShapeTypeMap.FromShapeType(st, auto, ph);
            }
            catch (Exception)
            {
            }

            if (existingType == "chart")
            {
                if (!PptHtmlChartIo.TryApplyToShape(
                    shape,
                    node.ChartGrid,
                    node.ChartFormat,
                    node.HasText,
                    warnings,
                    out error))
                {
                    return false;
                }
            }
            else if (existingType == "smartart" && node.HasText)
            {
                warnings.Add("忽略对 smartart 的文本修改: " + node.ShapeId);
            }
            else if ((existingType == "table" || node.TableCells != null) && node.TableCells != null)
            {
                if (!TryWriteTable(shape, node.TableCells, out error))
                {
                    return false;
                }
            }
            else if (node.HasText && existingType != "picture" && existingType != "media")
            {
                if (!TryWriteText(shape, node.Text ?? "", out error))
                {
                    return false;
                }
            }

            if (node.HasGeometry)
            {
                ApplyGeometry(shape, node, slideWidth, slideHeight);
            }

            if (node.Rotation.HasValue)
            {
                try
                {
                    shape.Rotation = (float)node.Rotation.Value;
                }
                catch (Exception)
                {
                }
            }

            // 先字体后颜色：写 NameFarEast 常会把主题字色重置为黑，字色必须最后落盘。
            if (!TryApplyFont(shape, node, existingType, out error))
            {
                return false;
            }

            if (!TryApplyParagraph(shape, node, existingType, out error))
            {
                return false;
            }

            if (!TryApplyLine(shape, node, existingType, out error))
            {
                return false;
            }

            if (!TryApplyColors(shape, node, existingType, out error))
            {
                return false;
            }

            // 字号/AutoSize 可能撑破形状：写完样式后重锁几何
            LockTextFrameAndGeometry(shape, node, slideWidth, slideHeight);

            // 换图，或 B2：页上仍是 freeform/smartart/group/unknown 时删旧 + AddPicture
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

                float left = shape.Left, top = shape.Top, width = shape.Width, height = shape.Height;
                try
                {
                    shape.Delete();
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
                if (!TryCreate(slide, node, slideWidth, slideHeight, warnings, out string newId, out error))
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

        private static bool TryCreate(
            PowerPoint.Slide slide,
            PptHtmlApplyNode node,
            float slideWidth,
            float slideHeight,
            List<string> warnings,
            out string newShapeId,
            out string error)
        {
            newShapeId = null;
            error = null;
            float left = (float)(node.LeftPct.GetValueOrDefault() / 100.0 * slideWidth);
            float top = (float)(node.TopPct.GetValueOrDefault() / 100.0 * slideHeight);
            float width = (float)(node.WidthPct.GetValueOrDefault() / 100.0 * slideWidth);
            float height = (float)(node.HeightPct.GetValueOrDefault() / 100.0 * slideHeight);
            if (width <= 0)
            {
                width = 100;
            }

            if (height <= 0)
            {
                height = 50;
            }

            PowerPoint.Shape shape;
            string type = node.ShapeType ?? "";
            try
            {
                if (type == "textbox")
                {
                    shape = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        left,
                        top,
                        width,
                        height);
                    if (node.HasText)
                    {
                        TryWriteText(shape, node.Text ?? "", out _);
                    }
                }
                else if (type == "table")
                {
                    int rows = Math.Max(1, node.TableCells == null ? 1 : node.TableCells.Count);
                    int cols = 1;
                    if (node.TableCells != null && node.TableCells.Count > 0)
                    {
                        cols = Math.Max(1, node.TableCells[0].Count);
                    }

                    if (rows > 20 || cols > 20)
                    {
                        error = "新建表格不得超过 20×20";
                        return false;
                    }

                    shape = slide.Shapes.AddTable(rows, cols, left, top, width, height);
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
                        shape = slide.Shapes.AddPicture(
                            picPath,
                            Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoTrue,
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
                else if (type == "chart")
                {
                    if (!PptHtmlChartIo.TryParseType(node.ChartType, out int xlType, out _, out error))
                    {
                        return false;
                    }

                    if (!PptHtmlChartIo.TryCreateOnSlide(
                        slide.Shapes,
                        left,
                        top,
                        width,
                        height,
                        xlType,
                        node.ChartGrid,
                        node.ChartFormat,
                        warnings,
                        out object created,
                        out error))
                    {
                        return false;
                    }

                    shape = (PowerPoint.Shape)created;
                }
                else if (type == "media")
                {
                    if (string.IsNullOrEmpty(node.ResolvedLocalPath) || !File.Exists(node.ResolvedLocalPath))
                    {
                        error = "媒体文件不存在: " + (node.DataSrc ?? "");
                        return false;
                    }

                    shape = slide.Shapes.AddMediaObject2(
                        node.ResolvedLocalPath,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        left,
                        top,
                        width,
                        height);
                }
                else if (PptShapeTypeMap.TryGetAutoShapeType(type, out int autoType) || type == "line")
                {
                    if (type == "line")
                    {
                        autoType = 9; // msoShapeOval fallback avoided; use straight line 1? 
                        // MsoAutoShapeType msoShapeLine / use AddLine
                        shape = slide.Shapes.AddLine(left, top, left + width, top + height);
                    }
                    else
                    {
                        shape = slide.Shapes.AddShape(
                            (Office.MsoAutoShapeType)autoType,
                            left,
                            top,
                            width,
                            height);
                        if (node.HasText && !string.IsNullOrEmpty(node.Text))
                        {
                            TryWriteText(shape, node.Text, out _);
                        }
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

            if (node.Rotation.HasValue)
            {
                try
                {
                    shape.Rotation = (float)node.Rotation.Value;
                }
                catch (Exception)
                {
                }
            }

            // 先字体后颜色：写 NameFarEast 常会把主题字色重置为黑，字色必须最后落盘。
            if (!TryApplyFont(shape, node, type, out error))
            {
                return false;
            }

            if (!TryApplyParagraph(shape, node, type, out error))
            {
                return false;
            }

            if (!TryApplyLine(shape, node, type, out error))
            {
                return false;
            }

            if (!TryApplyColors(shape, node, type, out error))
            {
                return false;
            }

            // 新建 AddTextbox 默认左右内边距约 7.2pt，会挤窄正文；无 data-margin-* 时置 0
            LockTextFrameAndGeometry(shape, node, slideWidth, slideHeight, zeroMarginsIfAbsent: true);

            try
            {
                int id = shape.Id;
                newShapeId = "sid" + planSlideId(slide) + "-s" + id.ToString(CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                newShapeId = "";
            }

            return true;
        }

        /// <summary>
        /// 关闭「形状随文字自动变大」，写文本框内边距，并按 HTML 百分比重锁几何，避免导航条等被撑开/盖住。
        /// </summary>
        private static void LockTextFrameAndGeometry(
            PowerPoint.Shape shape,
            PptHtmlApplyNode node,
            float slideWidth,
            float slideHeight)
        {
            LockTextFrameAndGeometry(shape, node, slideWidth, slideHeight, zeroMarginsIfAbsent: false);
        }

        private static void LockTextFrameAndGeometry(
            PowerPoint.Shape shape,
            PptHtmlApplyNode node,
            float slideWidth,
            float slideHeight,
            bool zeroMarginsIfAbsent)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                    ApplyTextMargins(shape.TextFrame, node, zeroMarginsIfAbsent);
                }
            }
            catch (Exception)
            {
            }

            if (node != null && node.HasGeometry)
            {
                ApplyGeometry(shape, node, slideWidth, slideHeight);
            }
        }

        private static void ApplyTextMargins(
            PowerPoint.TextFrame tf,
            PptHtmlApplyNode node,
            bool zeroIfAbsent)
        {
            if (tf == null || node == null)
            {
                return;
            }

            string type = node.ShapeType ?? "";
            if (type == "picture" || type == "media" || type == "table")
            {
                return;
            }

            bool any = node.MarginLeftPt.HasValue
                || node.MarginRightPt.HasValue
                || node.MarginTopPt.HasValue
                || node.MarginBottomPt.HasValue;
            if (!any && !zeroIfAbsent)
            {
                return;
            }

            try
            {
                if (node.MarginLeftPt.HasValue)
                {
                    tf.MarginLeft = (float)node.MarginLeftPt.Value;
                }
                else if (zeroIfAbsent)
                {
                    tf.MarginLeft = 0;
                }

                if (node.MarginRightPt.HasValue)
                {
                    tf.MarginRight = (float)node.MarginRightPt.Value;
                }
                else if (zeroIfAbsent)
                {
                    tf.MarginRight = 0;
                }

                if (node.MarginTopPt.HasValue)
                {
                    tf.MarginTop = (float)node.MarginTopPt.Value;
                }
                else if (zeroIfAbsent)
                {
                    tf.MarginTop = 0;
                }

                if (node.MarginBottomPt.HasValue)
                {
                    tf.MarginBottom = (float)node.MarginBottomPt.Value;
                }
                else if (zeroIfAbsent)
                {
                    tf.MarginBottom = 0;
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryApplyParagraph(
            PowerPoint.Shape shape,
            PptHtmlApplyNode node,
            string shapeType,
            out string error)
        {
            error = null;
            if (shape == null || node == null)
            {
                return true;
            }

            if (shapeType == "picture" || shapeType == "media" || shapeType == "table" || shapeType == "chart")
            {
                return true;
            }

            return PptHtmlParagraphIo.TryApplyToShape(
                shape,
                node.Align,
                node.LineSpacing,
                node.SpaceBeforePt,
                node.SpaceAfterPt,
                node.IndentLeftPt,
                node.IndentFirstPt,
                node.Bullet,
                out error);
        }

        private static bool TryApplyFont(
            PowerPoint.Shape shape,
            PptHtmlApplyNode node,
            string shapeType,
            out string error)
        {
            error = null;
            if (shape == null || node == null)
            {
                return true;
            }

            if (shapeType == "picture" || shapeType == "media" || shapeType == "table" || shapeType == "chart")
            {
                return true;
            }

            if (!node.FontSizePt.HasValue && !node.FontBold.HasValue && string.IsNullOrEmpty(node.FontName))
            {
                return true;
            }

            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return true;
                }

                PowerPoint.Font font = shape.TextFrame.TextRange.Font;
                if (node.FontSizePt.HasValue)
                {
                    font.Size = (float)node.FontSizePt.Value;
                }

                if (node.FontBold.HasValue)
                {
                    font.Bold = node.FontBold.Value
                        ? Office.MsoTriState.msoTrue
                        : Office.MsoTriState.msoFalse;
                }

                if (!string.IsNullOrEmpty(node.FontName))
                {
                    // 中文必须写 NameFarEast；仅 Name 时东亚字形仍可能是等线
                    try
                    {
                        font.NameFarEast = node.FontName;
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        font.NameAscii = node.FontName;
                    }
                    catch (Exception)
                    {
                    }

                    font.Name = node.FontName;
                }
            }
            catch (Exception ex)
            {
                error = "写字体样式失败: " + ex.Message;
                return false;
            }

            return true;
        }

        private static bool TryApplyLine(
            PowerPoint.Shape shape,
            PptHtmlApplyNode node,
            string shapeType,
            out string error)
        {
            error = null;
            if (shape == null || node == null)
            {
                return true;
            }

            if (shapeType == "picture" || shapeType == "media" || shapeType == "chart")
            {
                return true;
            }

            if (string.IsNullOrEmpty(node.LineColor) && !node.LineWidthPt.HasValue)
            {
                return true;
            }

            try
            {
                if (!string.IsNullOrEmpty(node.LineColor))
                {
                    if (string.Equals(node.LineColor, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        shape.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                    else
                    {
                        if (!PptHtmlStyleIo.TryParseHexToOfficeRgb(node.LineColor, out int rgb, out error))
                        {
                            return false;
                        }

                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.ForeColor.RGB = rgb;
                    }
                }

                if (node.LineWidthPt.HasValue)
                {
                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                    shape.Line.Weight = (float)node.LineWidthPt.Value;
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
            PowerPoint.Slide slide,
            PptHtmlApplyNode node,
            List<KeyValuePair<PowerPoint.Shape, int>> targets)
        {
            if (slide == null || node == null || !node.Z.HasValue || targets == null)
            {
                return;
            }

            if (!node.ShapeComId.HasValue)
            {
                return;
            }

            PowerPoint.Shape shape = FindShapeById(slide.Shapes, node.ShapeComId.Value);
            if (shape == null)
            {
                return;
            }

            targets.Add(new KeyValuePair<PowerPoint.Shape, int>(shape, node.Z.Value));
        }

        /// <summary>数值越大越靠上：按 z 升序依次 BringToFront。</summary>
        private static void ApplyZOrder(List<KeyValuePair<PowerPoint.Shape, int>> targets)
        {
            if (targets == null || targets.Count == 0)
            {
                return;
            }

            targets.Sort((a, b) => a.Value.CompareTo(b.Value));
            foreach (KeyValuePair<PowerPoint.Shape, int> item in targets)
            {
                try
                {
                    item.Key.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
                catch (Exception)
                {
                }
            }
        }

        private static void RelockAllGeometries(
            PowerPoint.Slide slide,
            List<PptHtmlApplyNode> nodes,
            float slideWidth,
            float slideHeight)
        {
            if (slide == null || nodes == null)
            {
                return;
            }

            foreach (PptHtmlApplyNode node in nodes)
            {
                if (node == null || !node.HasGeometry || !node.ShapeComId.HasValue)
                {
                    continue;
                }

                PowerPoint.Shape shape = FindShapeById(slide.Shapes, node.ShapeComId.Value);
                if (shape == null)
                {
                    continue;
                }

                LockTextFrameAndGeometry(shape, node, slideWidth, slideHeight);
            }
        }

        private static bool TryApplyColors(
            PowerPoint.Shape shape,
            PptHtmlApplyNode node,
            string shapeType,
            out string error)
        {
            error = null;
            if (shape == null || node == null)
            {
                return true;
            }

            if (shapeType == "picture" || shapeType == "media" || shapeType == "chart")
            {
                return true;
            }

            if (!string.IsNullOrEmpty(node.Fill))
            {
                try
                {
                    if (string.Equals(node.Fill, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        shape.Fill.Visible = Office.MsoTriState.msoFalse;
                    }
                    else
                    {
                        if (!PptHtmlStyleIo.TryParseHexToOfficeRgb(node.Fill, out int rgb, out error))
                        {
                            return false;
                        }

                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        try
                        {
                            shape.Fill.Solid();
                        }
                        catch (Exception)
                        {
                        }

                        shape.Fill.ForeColor.RGB = rgb;
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

                    bool wrote = false;
                    try
                    {
                        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb;
                        wrote = true;
                    }
                    catch (Exception)
                    {
                    }

                    if (!wrote && shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.Font.Color.RGB = rgb;
                    }
                }
                catch (Exception ex)
                {
                    error = "写字体颜色失败: " + ex.Message;
                    return false;
                }
            }

            return true;
        }

        private static string planSlideId(PowerPoint.Slide slide)
        {
            try
            {
                return Convert.ToString(slide.SlideID) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static int MapChartType(string chartType)
        {
            if (string.IsNullOrWhiteSpace(chartType))
            {
                return XlColumnClustered;
            }

            switch (chartType.Trim().ToLowerInvariant())
            {
                case "bar":
                    return XlBarClustered;
                case "line":
                    return XlLine;
                case "pie":
                    return XlPie;
                case "column":
                default:
                    return XlColumnClustered;
            }
        }

        private static void ApplyGeometry(
            PowerPoint.Shape shape,
            PptHtmlApplyNode node,
            float slideWidth,
            float slideHeight)
        {
            try
            {
                shape.Left = (float)(node.LeftPct.GetValueOrDefault() / 100.0 * slideWidth);
                shape.Top = (float)(node.TopPct.GetValueOrDefault() / 100.0 * slideHeight);
                shape.Width = (float)(node.WidthPct.GetValueOrDefault() / 100.0 * slideWidth);
                shape.Height = (float)(node.HeightPct.GetValueOrDefault() / 100.0 * slideHeight);
            }
            catch (Exception)
            {
            }
        }

        private static bool TryWriteText(PowerPoint.Shape shape, string text, out string error)
        {
            error = null;
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return true;
                }

                shape.TextFrame.TextRange.Text = text ?? "";
                return true;
            }
            catch (Exception ex)
            {
                error = "写文本失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryWriteTable(PowerPoint.Shape shape, List<List<string>> cells, out string error)
        {
            error = null;
            try
            {
                if (shape.HasTable != Office.MsoTriState.msoTrue)
                {
                    error = "目标不是表格";
                    return false;
                }

                PowerPoint.Table table = shape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                int wantRows = cells.Count;
                int wantCols = cells.Count == 0 ? 0 : cells[0].Count;
                if (wantRows > rows || wantCols > cols)
                {
                    error = "表格行列多于现有表（首期不自动扩表）";
                    return false;
                }

                for (int r = 0; r < wantRows; r++)
                {
                    List<string> row = cells[r];
                    for (int c = 0; c < row.Count && c < cols; c++)
                    {
                        table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Text = row[c] ?? "";
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

        private static PowerPoint.Shape FindShapeById(PowerPoint.Shapes shapes, int id)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape = shapes[i];
                    try
                    {
                        if (shape.Id == id)
                        {
                            return shape;
                        }

                        if ((int)shape.Type == 6)
                        {
                            PowerPoint.Shape found = FindInGroup(shape, id);
                            if (found != null)
                            {
                                return found;
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static PowerPoint.Shape FindInGroup(PowerPoint.Shape group, int id)
        {
            try
            {
                PowerPoint.GroupShapes items = group.GroupItems;
                int count = items.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape child = items[i];
                    if (child.Id == id)
                    {
                        return child;
                    }

                    if ((int)child.Type == 6)
                    {
                        PowerPoint.Shape nested = FindInGroup(child, id);
                        if (nested != null)
                        {
                            return nested;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
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
    }
}
