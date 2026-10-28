using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>幻灯片背景。PowerPoint 与 WPS 演示共用这一套，COM 走晚绑定。</summary>
    internal static class PptBackgroundEngine
    {
        private const int FillSolid = 1;
        private const int FillTextured = 4;
        private const int FillPicture = 6;

        public static bool TryRun(
            object presentation,
            string channelId,
            string hostKind,
            PresentationBackgroundRequest request,
            out PresentationBackgroundResult result,
            out string error)
        {
            result = null;
            error = null;
            if (presentation == null)
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "须提供 action（list|get|set|clear）";
                return false;
            }

            string action = request.Action.Trim().ToLowerInvariant();
            try
            {
                switch (action)
                {
                    case "list":
                        return TryList(presentation, channelId, hostKind, request, out result, out error);
                    case "get":
                        return TryGet(presentation, channelId, hostKind, request, out result, out error);
                    case "set":
                        return TrySet(presentation, channelId, hostKind, request, out result, out error);
                    case "clear":
                        return TryClear(presentation, channelId, hostKind, request, out result, out error);
                    default:
                        error = "action 须为 list|get|set|clear";
                        return false;
                }
            }
            catch (Exception ex)
            {
                error = "幻灯片背景失败: " + Unwrap(ex).Message;
                return false;
            }
        }

        private static bool TryList(
            object presentation,
            string channelId,
            string hostKind,
            PresentationBackgroundRequest request,
            out PresentationBackgroundResult result,
            out string error)
        {
            result = null;
            if (request.ListHasForbiddenFilter
                || request.ReadActionHasWriteFields
                || !string.IsNullOrWhiteSpace(request.SlideId))
            {
                error = "list 不接受 slide_id，也不接受改背景的字段";
                return false;
            }

            object slides = Prop(presentation, "Slides");
            int count = Convert.ToInt32(Prop(slides, "Count"));
            var list = new List<PresentationBackgroundInfo>();
            for (int i = 1; i <= count; i++)
            {
                object slide = WppCom.GetIndexed(slides, i);
                list.Add(Classify(slide, i));
            }

            result = Build(channelId, hostKind, "list", null, list);
            error = null;
            return true;
        }

        private static bool TryGet(
            object presentation,
            string channelId,
            string hostKind,
            PresentationBackgroundRequest request,
            out PresentationBackgroundResult result,
            out string error)
        {
            result = null;
            if (request.ReadActionHasWriteFields)
            {
                error = "get 不接受改背景的字段";
                return false;
            }

            if (!TryFindSlide(presentation, request.SlideId, out object slide, out int index, out string slideId, out error))
            {
                return false;
            }

            PresentationBackgroundInfo info = Classify(slide, index);
            if (string.Equals(info.Kind, "picture", StringComparison.Ordinal))
            {
                if (!TryReadLivePicture(presentation, index, request.ExportDirectory, out PptBackgroundPackage.PicturePlacement placement, out error))
                {
                    return false;
                }

                ApplyPlacement(info, placement, placement.RelativePath);
            }

            result = Build(channelId, hostKind, "get", slideId, new List<PresentationBackgroundInfo> { info });
            return true;
        }

        private static bool TryClear(
            object presentation,
            string channelId,
            string hostKind,
            PresentationBackgroundRequest request,
            out PresentationBackgroundResult result,
            out string error)
        {
            result = null;
            if (request.ReadActionHasWriteFields)
            {
                error = "clear 不接受颜色、图片或摆放字段";
                return false;
            }

            if (!TryFindSlide(presentation, request.SlideId, out object slide, out int index, out string slideId, out error))
            {
                return false;
            }

            string backup = PptHtmlOoxmlIo.TrySaveCopyTemp(presentation, "bg-clear", out error);
            if (string.IsNullOrEmpty(backup))
            {
                return false;
            }

            try
            {
                SetProp(slide, "FollowMasterBackground", -1);
                PresentationBackgroundInfo info = Classify(slide, index);
                if (!info.FollowMaster)
                {
                    error = "clear 后这一页仍未跟随母版";
                    TryRestore(presentation, index, backup);
                    return false;
                }

                result = Build(channelId, hostKind, "clear", slideId, new List<PresentationBackgroundInfo> { info });
                return true;
            }
            catch (Exception ex)
            {
                TryRestore(presentation, index, backup);
                error = "恢复跟随母版失败: " + Unwrap(ex).Message;
                return false;
            }
            finally
            {
                PptHtmlOoxmlIo.TryDeleteFile(backup);
            }
        }

        private static bool TrySet(
            object presentation,
            string channelId,
            string hostKind,
            PresentationBackgroundRequest request,
            out PresentationBackgroundResult result,
            out string error)
        {
            result = null;
            if (!TryNormalizeSet(request, out string kind, out int bgr, out PptBackgroundPackage.PicturePlacement placement, out error))
            {
                return false;
            }

            if (!TryFindSlide(presentation, request.SlideId, out object slide, out int index, out string slideId, out error))
            {
                return false;
            }

            string backup = PptHtmlOoxmlIo.TrySaveCopyTemp(presentation, "bg-set", out error);
            if (string.IsNullOrEmpty(backup))
            {
                return false;
            }

            try
            {
                SetProp(slide, "FollowMasterBackground", 0);
                object fill = FillOf(slide);
                if (string.Equals(kind, "solid", StringComparison.Ordinal))
                {
                    Call(fill, "Solid");
                    SetProp(Prop(fill, "ForeColor"), "RGB", bgr);
                    PresentationBackgroundInfo solidInfo = Classify(slide, index);
                    if (!string.Equals(solidInfo.Kind, "solid", StringComparison.Ordinal)
                        || !string.Equals(solidInfo.Color, request.Color.Trim().ToUpperInvariant(), StringComparison.Ordinal))
                    {
                        error = "纯色写回后读到的颜色不一致";
                        TryRestore(presentation, index, backup);
                        return false;
                    }

                    result = Build(channelId, hostKind, "set", slideId, new List<PresentationBackgroundInfo> { solidInfo });
                    return true;
                }

                Call(fill, "UserPicture", request.PictureAbsolutePath);
                bool tile = string.Equals(placement.Fit, "tile", StringComparison.Ordinal);
                SetProp(fill, "TextureTile", tile ? -1 : 0);
                if (tile)
                {
                    SetProp(fill, "TextureOffsetX", placement.TileOffsetX);
                    SetProp(fill, "TextureOffsetY", placement.TileOffsetY);
                    SetProp(fill, "TextureHorizontalScale", placement.TileScaleX);
                    SetProp(fill, "TextureVerticalScale", placement.TileScaleY);
                    SetProp(fill, "TextureAlignment", AlignToCom(placement.TileAlign));
                }

                bool needXml = tile
                    ? !string.Equals(placement.TileMirror, "none", StringComparison.Ordinal)
                    : !AllStretchZero(placement);
                if (!needXml && !LivePictureMatches(presentation, index, placement, out _))
                {
                    needXml = true;
                }

                if (needXml)
                {
                    if (!TryPatchLive(presentation, index, placement, out error))
                    {
                        TryRestore(presentation, index, backup);
                        return false;
                    }
                }

                SetProp(slide, "FollowMasterBackground", 0);
                if (!LivePictureMatches(presentation, index, placement, out string mismatch))
                {
                    error = "图片背景写回后摆放不一致: " + mismatch;
                    TryRestore(presentation, index, backup);
                    return false;
                }

                PresentationBackgroundInfo info = Classify(slide, index);
                ApplyPlacement(info, placement, request.PathEcho);
                result = Build(channelId, hostKind, "set", slideId, new List<PresentationBackgroundInfo> { info });
                return true;
            }
            catch (Exception ex)
            {
                TryRestore(presentation, index, backup);
                error = "设置背景失败: " + Unwrap(ex).Message + "。已尝试回滚，若窗口里的背景变了，说明回滚未完成";
                return false;
            }
            finally
            {
                PptHtmlOoxmlIo.TryDeleteFile(backup);
            }
        }

        private static bool TryPatchLive(
            object presentation,
            int slideIndex,
            PptBackgroundPackage.PicturePlacement placement,
            out string error)
        {
            string edited = PptHtmlOoxmlIo.TrySaveCopyTemp(presentation, "bg-patch", out error);
            if (string.IsNullOrEmpty(edited))
            {
                return false;
            }

            try
            {
                if (!PptBackgroundPackage.TryRewritePlacement(edited, slideIndex, placement, out error))
                {
                    return false;
                }

                if (!TryPickUpBackground(presentation, slideIndex, edited, out error))
                {
                    return false;
                }

                return true;
            }
            finally
            {
                PptHtmlOoxmlIo.TryDeleteFile(edited);
            }
        }

        private static bool TryPickUpBackground(
            object livePresentation,
            int slideIndex,
            string packagePath,
            out string error)
        {
            error = null;
            object app;
            try
            {
                app = Prop(livePresentation, "Application");
            }
            catch (Exception ex)
            {
                error = "贴回背景：拿不到宿主 " + Unwrap(ex).Message;
                return false;
            }

            object copy = OpenHidden(app, packagePath, out error);
            if (copy == null)
            {
                return false;
            }

            try
            {
                object srcSlide = SlideAt(copy, slideIndex);
                object destSlide = SlideAt(livePresentation, slideIndex);
                object srcBg = Prop(srcSlide, "Background");
                object destBg = Prop(destSlide, "Background");
                Call(srcBg, "PickUp");
                Call(destBg, "Apply");
                SetProp(destSlide, "FollowMasterBackground", 0);
                return true;
            }
            catch (Exception ex)
            {
                error = "贴回背景失败: " + Unwrap(ex).Message;
                return false;
            }
            finally
            {
                CloseCopy(copy);
            }
        }

        private static void TryRestore(object livePresentation, int slideIndex, string backupPath)
        {
            if (string.IsNullOrEmpty(backupPath) || livePresentation == null)
            {
                return;
            }

            try
            {
                TryPickUpBackground(livePresentation, slideIndex, backupPath, out _);
            }
            catch (Exception)
            {
            }
        }

        private static bool LivePictureMatches(
            object presentation,
            int slideIndex,
            PptBackgroundPackage.PicturePlacement expected,
            out string detail)
        {
            detail = null;
            if (!TryReadLivePicture(presentation, slideIndex, null, out PptBackgroundPackage.PicturePlacement actual, out string error))
            {
                detail = error ?? "读回失败";
                return false;
            }

            return PptBackgroundPackage.Matches(actual, expected, out detail);
        }

        private static bool TryReadLivePicture(
            object presentation,
            int slideIndex,
            string exportDirectory,
            out PptBackgroundPackage.PicturePlacement placement,
            out string error)
        {
            placement = null;
            PptBackgroundPackage.PicturePlacement read = null;
            string readError = null;
            bool ok = PptHtmlOoxmlIo.TryReadSavedPackage(
                presentation,
                path => PptBackgroundPackage.TryReadPicture(
                    path,
                    slideIndex,
                    exportDirectory,
                    out read,
                    out readError),
                out error,
                "背景图");
            if (!ok)
            {
                if (!string.IsNullOrEmpty(readError))
                {
                    error = readError;
                }

                return false;
            }

            placement = read;
            error = null;
            return placement != null;
        }

        private static PresentationBackgroundInfo Classify(object slide, int fallbackIndex)
        {
            var info = new PresentationBackgroundInfo
            {
                SlideId = "",
                Index = fallbackIndex,
                Kind = "other",
                FollowMaster = false
            };
            try
            {
                info.SlideId = Convert.ToString(Prop(slide, "SlideID")) ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                info.Index = Convert.ToInt32(Prop(slide, "SlideIndex"));
            }
            catch (Exception)
            {
            }

            if (IsTrue(Prop(slide, "FollowMasterBackground")))
            {
                info.FollowMaster = true;
                info.Kind = "master";
                return info;
            }

            object fill = FillOf(slide);
            int fillType = Convert.ToInt32(Prop(fill, "Type"));
            if (fillType == FillSolid)
            {
                info.Kind = "solid";
                info.Color = RgbToHex(Convert.ToInt32(Prop(Prop(fill, "ForeColor"), "RGB")));
                return info;
            }

            if (fillType == FillPicture || (fillType == FillTextured && IsUserTexture(fill)))
            {
                info.Kind = "picture";
                return info;
            }

            info.Kind = "other";
            return info;
        }

        private static bool IsUserTexture(object fill)
        {
            try
            {
                int textureType = Convert.ToInt32(Prop(fill, "TextureType"));
                return textureType == (int)Office.MsoTextureType.msoTextureUserDefined;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void ApplyPlacement(
            PresentationBackgroundInfo info,
            PptBackgroundPackage.PicturePlacement placement,
            string path)
        {
            if (info == null || placement == null)
            {
                return;
            }

            info.Kind = "picture";
            info.FollowMaster = false;
            info.Path = path;
            info.Fit = placement.Fit;
            if (string.Equals(placement.Fit, "tile", StringComparison.Ordinal))
            {
                info.TileOffsetX = placement.TileOffsetX;
                info.TileOffsetY = placement.TileOffsetY;
                info.TileScaleX = placement.TileScaleX;
                info.TileScaleY = placement.TileScaleY;
                info.TileAlign = placement.TileAlign;
                info.TileMirror = placement.TileMirror;
                return;
            }

            info.StretchLeft = placement.StretchLeft;
            info.StretchRight = placement.StretchRight;
            info.StretchTop = placement.StretchTop;
            info.StretchBottom = placement.StretchBottom;
        }

        private static bool AllStretchZero(PptBackgroundPackage.PicturePlacement placement)
        {
            return Math.Abs(placement.StretchLeft) <= 0.1
                && Math.Abs(placement.StretchRight) <= 0.1
                && Math.Abs(placement.StretchTop) <= 0.1
                && Math.Abs(placement.StretchBottom) <= 0.1;
        }

        private static bool TryNormalizeSet(
            PresentationBackgroundRequest request,
            out string kind,
            out int bgr,
            out PptBackgroundPackage.PicturePlacement placement,
            out string error)
        {
            kind = (request.Kind ?? "").Trim().ToLowerInvariant();
            bgr = 0;
            placement = null;
            error = null;
            if (kind != "solid" && kind != "picture")
            {
                error = "set 的 kind 须为 solid 或 picture";
                return false;
            }

            bool hasColor = !string.IsNullOrWhiteSpace(request.Color);
            bool hasPath = !string.IsNullOrWhiteSpace(request.PictureAbsolutePath);
            if (kind == "solid")
            {
                if (!hasColor || hasPath || request.SawStretch || request.SawTile || !string.IsNullOrEmpty(request.Fit))
                {
                    error = "solid 只接受 color，不接受 path 或摆放字段";
                    return false;
                }

                if (!TryParseHex(request.Color, out bgr, out string hex, out error))
                {
                    return false;
                }

                request.Color = hex;
                return true;
            }

            if (!hasPath || hasColor)
            {
                error = "picture 必须提供 path，不接受 color";
                return false;
            }

            if (!File.Exists(request.PictureAbsolutePath))
            {
                error = "图片不存在: " + request.PictureAbsolutePath;
                return false;
            }

            return TryBuildPlacement(request, out placement, out error);
        }

        private static bool TryBuildPlacement(
            PresentationBackgroundRequest request,
            out PptBackgroundPackage.PicturePlacement placement,
            out string error)
        {
            placement = null;
            error = null;
            string fit = string.IsNullOrWhiteSpace(request.Fit) ? null : request.Fit.Trim().ToLowerInvariant();
            if (fit == null)
            {
                if (request.SawStretch && request.SawTile)
                {
                    error = "拉伸字段和平铺字段不能一起传";
                    return false;
                }

                fit = request.SawTile ? "tile" : "stretch";
            }
            else if (fit != "stretch" && fit != "tile")
            {
                error = "fit 须为 stretch 或 tile";
                return false;
            }

            if (fit == "stretch" && request.SawTile)
            {
                error = "fit=stretch 不能带平铺字段";
                return false;
            }

            if (fit == "tile" && request.SawStretch)
            {
                error = "fit=tile 不能带拉伸字段";
                return false;
            }

            if (fit == "stretch")
            {
                placement = new PptBackgroundPackage.PicturePlacement
                {
                    Fit = "stretch",
                    StretchLeft = request.StretchLeft ?? 0,
                    StretchRight = request.StretchRight ?? 0,
                    StretchTop = request.StretchTop ?? 0,
                    StretchBottom = request.StretchBottom ?? 0
                };
                return true;
            }

            string align = string.IsNullOrWhiteSpace(request.TileAlign) ? "tl" : request.TileAlign.Trim().ToLowerInvariant();
            string mirror = string.IsNullOrWhiteSpace(request.TileMirror) ? "none" : request.TileMirror.Trim().ToLowerInvariant();
            if (!IsAlign(align))
            {
                error = "tile_align 须为 tl/t/tr/l/ctr/r/bl/b/br";
                return false;
            }

            if (mirror != "none" && mirror != "x" && mirror != "y" && mirror != "xy")
            {
                error = "tile_mirror 须为 none/x/y/xy";
                return false;
            }

            double sx = request.TileScaleX ?? 100;
            double sy = request.TileScaleY ?? 100;
            if (sx <= 0 || sy <= 0)
            {
                error = "平铺缩放必须大于 0";
                return false;
            }

            placement = new PptBackgroundPackage.PicturePlacement
            {
                Fit = "tile",
                TileOffsetX = request.TileOffsetX ?? 0,
                TileOffsetY = request.TileOffsetY ?? 0,
                TileScaleX = sx,
                TileScaleY = sy,
                TileAlign = align,
                TileMirror = mirror
            };
            return true;
        }

        private static bool IsAlign(string align)
        {
            switch (align)
            {
                case "tl":
                case "t":
                case "tr":
                case "l":
                case "ctr":
                case "r":
                case "bl":
                case "b":
                case "br":
                    return true;
                default:
                    return false;
            }
        }

        private static int AlignToCom(string align)
        {
            Office.MsoTextureAlignment value;
            switch (align)
            {
                case "t":
                    value = Office.MsoTextureAlignment.msoTextureTop;
                    break;
                case "tr":
                    value = Office.MsoTextureAlignment.msoTextureTopRight;
                    break;
                case "l":
                    value = Office.MsoTextureAlignment.msoTextureLeft;
                    break;
                case "ctr":
                    value = Office.MsoTextureAlignment.msoTextureCenter;
                    break;
                case "r":
                    value = Office.MsoTextureAlignment.msoTextureRight;
                    break;
                case "bl":
                    value = Office.MsoTextureAlignment.msoTextureBottomLeft;
                    break;
                case "b":
                    value = Office.MsoTextureAlignment.msoTextureBottom;
                    break;
                case "br":
                    value = Office.MsoTextureAlignment.msoTextureBottomRight;
                    break;
                default:
                    value = Office.MsoTextureAlignment.msoTextureTopLeft;
                    break;
            }

            return (int)value;
        }

        private static bool TryParseHex(string raw, out int bgr, out string hex, out string error)
        {
            bgr = 0;
            hex = null;
            error = null;
            string text = (raw ?? "").Trim();
            if (text.Length != 7 || text[0] != '#')
            {
                error = "color 须为 #RRGGBB";
                return false;
            }

            int r;
            int g;
            int b;
            if (!int.TryParse(text.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r)
                || !int.TryParse(text.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g)
                || !int.TryParse(text.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b))
            {
                error = "color 须为 #RRGGBB";
                return false;
            }

            bgr = r | (g << 8) | (b << 16);
            hex = "#" + r.ToString("X2", CultureInfo.InvariantCulture)
                + g.ToString("X2", CultureInfo.InvariantCulture)
                + b.ToString("X2", CultureInfo.InvariantCulture);
            return true;
        }

        private static string RgbToHex(int officeRgb)
        {
            int r = officeRgb & 0xFF;
            int g = (officeRgb >> 8) & 0xFF;
            int b = (officeRgb >> 16) & 0xFF;
            return "#" + r.ToString("X2", CultureInfo.InvariantCulture)
                + g.ToString("X2", CultureInfo.InvariantCulture)
                + b.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static bool TryFindSlide(
            object presentation,
            string slideIdRaw,
            out object slide,
            out int index,
            out string slideId,
            out string error)
        {
            slide = null;
            index = 0;
            slideId = null;
            error = null;
            if (string.IsNullOrWhiteSpace(slideIdRaw))
            {
                error = "须提供 slide_id";
                return false;
            }

            slideId = slideIdRaw.Trim();
            if (slideId.StartsWith("sid", StringComparison.OrdinalIgnoreCase)
                || !int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idInt))
            {
                error = "非法 slide_id: " + slideId + "（抄概览里的 slide_id，不要加 sid 前缀，也不要用页码）";
                return false;
            }

            object slides = Prop(presentation, "Slides");
            int count = Convert.ToInt32(Prop(slides, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object candidate = WppCom.GetIndexed(slides, i);
                int candidateId = Convert.ToInt32(Prop(candidate, "SlideID"));
                if (candidateId == idInt)
                {
                    slide = candidate;
                    index = i;
                    return true;
                }
            }

            error = "幻灯片不存在: slide_id=" + slideId;
            return false;
        }

        private static object SlideAt(object presentation, int index)
        {
            object slides = Prop(presentation, "Slides");
            object slide = WppCom.GetIndexed(slides, index);
            if (slide == null)
            {
                throw new InvalidOperationException("找不到第 " + index + " 页");
            }

            return slide;
        }

        private static object FillOf(object slide)
        {
            object background = Prop(slide, "Background");
            object fill = Prop(background, "Fill");
            if (fill == null)
            {
                throw new InvalidOperationException("这一页没有 Fill");
            }

            return fill;
        }

        private static object OpenHidden(object app, string path, out string error)
        {
            error = null;
            PowerPoint.Application pptApp = app as PowerPoint.Application;
            if (pptApp == null)
            {
                try
                {
                    object maybe = Prop(app, "Presentations");
                    if (maybe != null && app != null)
                    {
                        PowerPoint.Presentation probe = app as PowerPoint.Presentation;
                        if (probe != null)
                        {
                            pptApp = probe.Application;
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            if (pptApp != null)
            {
                try
                {
                    return pptApp.Presentations.Open(
                        path,
                        Office.MsoTriState.msoTrue,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoFalse);
                }
                catch (Exception ex)
                {
                    error = "打开背景副本失败: " + Unwrap(ex).Message;
                    return null;
                }
            }

            object presentations;
            try
            {
                presentations = Prop(app, "Presentations");
            }
            catch (Exception ex)
            {
                error = "打开背景副本失败: " + Unwrap(ex).Message;
                return null;
            }

            try
            {
                object opened = Call(presentations, "Open", path, true, false, false);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object opened = Call(presentations, "Open", path, true);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception ex)
            {
                error = "打开背景副本失败: " + Unwrap(ex).Message;
                return null;
            }

            error = "打开背景副本无返回";
            return null;
        }

        private static void CloseCopy(object copy)
        {
            if (copy == null)
            {
                return;
            }

            try
            {
                SetProp(copy, "Saved", -1);
            }
            catch (Exception)
            {
            }

            try
            {
                Call(copy, "Close");
            }
            catch (Exception)
            {
            }
        }

        private static PresentationBackgroundResult Build(
            string channelId,
            string hostKind,
            string action,
            string slideId,
            List<PresentationBackgroundInfo> backgrounds)
        {
            return new PresentationBackgroundResult
            {
                ChannelId = channelId ?? "",
                Kind = hostKind ?? "",
                Action = action,
                SlideId = slideId ?? "",
                Backgrounds = backgrounds ?? new List<PresentationBackgroundInfo>()
            };
        }

        private static bool IsTrue(object value)
        {
            if (value == null)
            {
                return false;
            }

            if (value is bool b)
            {
                return b;
            }

            try
            {
                return Convert.ToInt32(value) != 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static object Prop(object target, string name)
        {
            try
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    target,
                    null);
            }
            catch (TargetInvocationException ex)
            {
                throw Unwrap(ex);
            }
        }

        private static void SetProp(object target, string name, object value)
        {
            try
            {
                target.GetType().InvokeMember(
                    name,
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    target,
                    new[] { value });
            }
            catch (TargetInvocationException ex)
            {
                throw Unwrap(ex);
            }
        }

        private static object Call(object target, string name, params object[] args)
        {
            try
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    target,
                    args == null || args.Length == 0 ? null : args);
            }
            catch (TargetInvocationException ex)
            {
                throw Unwrap(ex);
            }
        }

        private static Exception Unwrap(Exception ex)
        {
            while (ex is TargetInvocationException tie && tie.InnerException != null)
            {
                ex = tie.InnerException;
            }

            return ex;
        }
    }
}
