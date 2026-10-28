using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_PptBackgroundTool
    {
        private static readonly string[] StretchKeys =
        {
            "stretch_left", "stretch_right", "stretch_top", "stretch_bottom"
        };

        private static readonly string[] TileKeys =
        {
            "tile_offset_x", "tile_offset_y", "tile_scale_x", "tile_scale_y", "tile_align", "tile_mirror"
        };

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_ppt_background"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!TryBuildRequest(args, out PresentationBackgroundRequest request, out string buildError))
                    {
                        return new ToolResult { Success = false, Error = buildError };
                    }

                    if (string.Equals(request.Kind, "picture", StringComparison.OrdinalIgnoreCase)
                        && HasArg(args, "path"))
                    {
                        string pathArg = GetStringArg(args, "path");
                        (bool pathOk, string absolute, string pathError) = await TryResolvePicturePathAsync(pathArg).ConfigureAwait(false);
                        if (!pathOk)
                        {
                            return new ToolResult { Success = false, Error = pathError };
                        }

                        request.PictureAbsolutePath = absolute;
                        request.PathEcho = pathArg;
                    }

                    if (!PresentationHostAdapter.TryPptBackground(
                            channel,
                            request,
                            out PresentationBackgroundResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var backgrounds = new List<Dictionary<string, object>>();
                    if (hostResult.Backgrounds != null)
                    {
                        foreach (PresentationBackgroundInfo info in hostResult.Backgrounds)
                        {
                            if (info != null)
                            {
                                backgrounds.Add(InfoToDict(info));
                            }
                        }
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["backgrounds"] = backgrounds,
                        ["display_contents"] = BuildDisplayContents(hostResult)
                    };
                    if (!string.IsNullOrEmpty(hostResult.SlideId))
                    {
                        data["slide_id"] = hostResult.SlideId;
                    }

                    await Task.CompletedTask;
                    return new ToolResult { Success = true, Data = data };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "幻灯片背景失败: " + ex.Message };
                }
            };
        }

        public static bool TryGetContractParseFailure(Dictionary<string, object> args, out ToolResult failure)
        {
            failure = null;
            if (string.IsNullOrWhiteSpace(GetStringArg(args, "action")))
            {
                failure = new ToolResult { Success = false, Error = "须提供 action（list|get|set|clear）" };
                return true;
            }

            return false;
        }

        private static bool TryBuildRequest(
            Dictionary<string, object> args,
            out PresentationBackgroundRequest request,
            out string error)
        {
            request = null;
            error = null;
            string action = GetStringArg(args, "action");
            if (string.IsNullOrWhiteSpace(action))
            {
                error = "须提供 action（list|get|set|clear）";
                return false;
            }

            string actionNorm = action.Trim().ToLowerInvariant();
            bool sawStretch = HasAny(args, StretchKeys);
            bool sawTile = HasAny(args, TileKeys);
            bool sawWrite = sawStretch || sawTile || HasArg(args, "kind") || HasArg(args, "color")
                || HasArg(args, "path") || HasArg(args, "fit");
            string exportDir = null;
            if (string.Equals(actionNorm, "get", StringComparison.Ordinal))
            {
                exportDir = Path.Combine(
                    WorkspacePathResolver.GetSessionDirectory(),
                    PptBackgroundPackage.ExportFolderName);
            }

            double? stretchLeft;
            double? stretchRight;
            double? stretchTop;
            double? stretchBottom;
            double? tileOffsetX;
            double? tileOffsetY;
            double? tileScaleX;
            double? tileScaleY;
            if (!TryOptionalNumber(args, "stretch_left", out stretchLeft, out error)
                || !TryOptionalNumber(args, "stretch_right", out stretchRight, out error)
                || !TryOptionalNumber(args, "stretch_top", out stretchTop, out error)
                || !TryOptionalNumber(args, "stretch_bottom", out stretchBottom, out error)
                || !TryOptionalNumber(args, "tile_offset_x", out tileOffsetX, out error)
                || !TryOptionalNumber(args, "tile_offset_y", out tileOffsetY, out error)
                || !TryOptionalNumber(args, "tile_scale_x", out tileScaleX, out error)
                || !TryOptionalNumber(args, "tile_scale_y", out tileScaleY, out error))
            {
                return false;
            }

            request = new PresentationBackgroundRequest
            {
                Action = action,
                SlideId = GetStringArg(args, "slide_id"),
                ListHasForbiddenFilter = HasArg(args, "slide_ids")
                    || (string.Equals(actionNorm, "list", StringComparison.Ordinal) && HasArg(args, "slide_id")),
                ReadActionHasWriteFields = !string.Equals(actionNorm, "set", StringComparison.Ordinal) && sawWrite,
                Kind = GetStringArg(args, "kind"),
                Color = GetStringArg(args, "color"),
                PathEcho = GetStringArg(args, "path"),
                Fit = GetStringArg(args, "fit"),
                SawStretch = sawStretch,
                SawTile = sawTile,
                StretchLeft = stretchLeft,
                StretchRight = stretchRight,
                StretchTop = stretchTop,
                StretchBottom = stretchBottom,
                TileOffsetX = tileOffsetX,
                TileOffsetY = tileOffsetY,
                TileScaleX = tileScaleX,
                TileScaleY = tileScaleY,
                TileAlign = GetStringArg(args, "tile_align"),
                TileMirror = GetStringArg(args, "tile_mirror"),
                ExportDirectory = exportDir
            };
            return true;
        }

        private static async Task<(bool ok, string absolute, string error)> TryResolvePicturePathAsync(string pathArg)
        {
            if (!PptHtmlFileSource.TryClassify(
                    pathArg,
                    out bool isAbsolute,
                    out bool isWorkspace,
                    out string workspaceName,
                    out string classifyError))
            {
                return (false, null, classifyError);
            }

            if (isAbsolute)
            {
                string path = Path.GetFullPath(pathArg.Trim());
                if (!File.Exists(path))
                {
                    return (false, null, "文件不存在: " + path);
                }

                return (true, path, null);
            }

            if (!isWorkspace)
            {
                return (false, null, "图片路径无效");
            }

            if (FilePathResolver.TryResolve(workspaceName, out ResolvedFilePath resolved, out _)
                && !string.IsNullOrEmpty(resolved.LocalPath)
                && File.Exists(resolved.LocalPath))
            {
                return (true, Path.GetFullPath(resolved.LocalPath), null);
            }

            if (!FilePathResolver.TryResolve(workspaceName, out resolved, out string picError))
            {
                return (false, null, "工作区图片路径无效: " + workspaceName + "（" + picError + "）");
            }

            var picRead = await FilePathResolver.ReadBytesAsync(resolved).ConfigureAwait(false);
            if (!picRead.Success || string.IsNullOrEmpty(resolved.LocalPath) || !File.Exists(resolved.LocalPath))
            {
                return (false, null, "工作区图片不可用: " + workspaceName + "（" + (picRead.Error ?? "本地与云端均未找到") + "）");
            }

            return (true, Path.GetFullPath(resolved.LocalPath), null);
        }

        private static Dictionary<string, object> InfoToDict(PresentationBackgroundInfo info)
        {
            var d = new Dictionary<string, object>
            {
                ["slide_id"] = info.SlideId ?? "",
                ["index"] = info.Index,
                ["follow_master"] = info.FollowMaster,
                ["kind"] = info.Kind ?? ""
            };
            if (!string.IsNullOrEmpty(info.Color))
            {
                d["color"] = info.Color;
            }

            if (!string.IsNullOrEmpty(info.Path))
            {
                d["path"] = info.Path;
            }

            if (!string.IsNullOrEmpty(info.Fit))
            {
                d["fit"] = info.Fit;
            }

            AddNumber(d, "stretch_left", info.StretchLeft);
            AddNumber(d, "stretch_right", info.StretchRight);
            AddNumber(d, "stretch_top", info.StretchTop);
            AddNumber(d, "stretch_bottom", info.StretchBottom);
            AddNumber(d, "tile_offset_x", info.TileOffsetX);
            AddNumber(d, "tile_offset_y", info.TileOffsetY);
            AddNumber(d, "tile_scale_x", info.TileScaleX);
            AddNumber(d, "tile_scale_y", info.TileScaleY);
            if (!string.IsNullOrEmpty(info.TileAlign))
            {
                d["tile_align"] = info.TileAlign;
            }

            if (!string.IsNullOrEmpty(info.TileMirror))
            {
                d["tile_mirror"] = info.TileMirror;
            }

            return d;
        }

        private static void AddNumber(Dictionary<string, object> d, string key, double? value)
        {
            if (!value.HasValue)
            {
                return;
            }

            d[key] = value.Value;
        }

        private static string BuildDisplayContents(PresentationBackgroundResult result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("| slide_id | index | follow_master | kind | color | fit | path |");
            sb.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");
            if (result?.Backgrounds != null)
            {
                foreach (PresentationBackgroundInfo info in result.Backgrounds)
                {
                    if (info == null)
                    {
                        continue;
                    }

                    sb.Append("| ")
                        .Append(info.SlideId ?? "")
                        .Append(" | ")
                        .Append(info.Index)
                        .Append(" | ")
                        .Append(info.FollowMaster ? "true" : "false")
                        .Append(" | ")
                        .Append(info.Kind ?? "")
                        .Append(" | ")
                        .Append(info.Color ?? "")
                        .Append(" | ")
                        .Append(info.Fit ?? "")
                        .Append(" | ")
                        .Append(info.Path ?? "")
                        .AppendLine(" |");
                }
            }

            sb.AppendLine();
            sb.AppendLine("说明：幻灯片背景挂靠 slide_id。list 不带图片。get 才导出这一页自己的背景原图。铺满的图片形状不是背景。只传 path 会变成拉伸贴边。纯色再写回是固定色。");
            return sb.ToString();
        }

        private static bool HasAny(Dictionary<string, object> args, string[] keys)
        {
            if (keys == null)
            {
                return false;
            }

            foreach (string key in keys)
            {
                if (HasArg(args, key))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HasArg(Dictionary<string, object> args, string key)
        {
            if (args == null || string.IsNullOrEmpty(key))
            {
                return false;
            }

            foreach (KeyValuePair<string, object> kv in args)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    return true;
                }
            }

            return false;
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || string.IsNullOrEmpty(key))
            {
                return null;
            }

            foreach (KeyValuePair<string, object> kv in args)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    string s = Convert.ToString(kv.Value);
                    return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
                }
            }

            return null;
        }

        private static bool TryOptionalNumber(
            Dictionary<string, object> args,
            string key,
            out double? value,
            out string error)
        {
            value = null;
            error = null;
            if (!HasArg(args, key))
            {
                return true;
            }

            object raw = null;
            foreach (KeyValuePair<string, object> kv in args)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase))
                {
                    raw = kv.Value;
                    break;
                }
            }

            if (raw == null)
            {
                return true;
            }

            if (raw is double d)
            {
                value = d;
                return true;
            }

            if (raw is float f)
            {
                value = f;
                return true;
            }

            if (raw is int i)
            {
                value = i;
                return true;
            }

            if (raw is long l)
            {
                value = l;
                return true;
            }

            if (double.TryParse(Convert.ToString(raw, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                || double.TryParse(Convert.ToString(raw), NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
            {
                value = parsed;
                return true;
            }

            error = key + " 须为数字";
            return false;
        }
    }
}
