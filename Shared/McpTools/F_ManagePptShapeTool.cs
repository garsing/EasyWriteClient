using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_ManagePptShapeTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_manage_ppt_shape"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel dest, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    string action = GetStringArg(args, "action");
                    if (string.IsNullOrWhiteSpace(action))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "须提供 action（delete、group 或 duplicate_group）"
                        };
                    }

                    action = action.Trim().ToLowerInvariant();
                    if (!TryRejectForeignArgs(action, args, out string foreignError))
                    {
                        return new ToolResult { Success = false, Error = foreignError };
                    }

                    if (!TryResolveSourceChannel(action, dest, args, out IOperationChannel source, out string sourceError))
                    {
                        return new ToolResult { Success = false, Error = sourceError };
                    }

                    var request = new PresentationManageShapeRequest
                    {
                        Action = action,
                        SlideId = GetStringArg(args, "slide_id"),
                        ToSlideId = GetStringArg(args, "to_slide_id"),
                        ShapeIds = TryGetStringArray(args, "shape_ids"),
                        ShapeId = GetStringArg(args, "shape_id"),
                        SourceChannelId = source.ChannelId
                    };

                    if (string.Equals(action, "delete", StringComparison.Ordinal)
                        && (request.ShapeIds == null || request.ShapeIds.Length == 0)
                        && !string.IsNullOrWhiteSpace(request.ShapeId))
                    {
                        request.ShapeIds = new[] { request.ShapeId };
                    }

                    if (string.Equals(action, "duplicate_group", StringComparison.Ordinal))
                    {
                        if (!TryParsePct(args, "left", out double left, out string leftError))
                        {
                            return new ToolResult { Success = false, Error = leftError };
                        }

                        if (!TryParsePct(args, "top", out double top, out string topError))
                        {
                            return new ToolResult { Success = false, Error = topError };
                        }

                        request.LeftPct = left;
                        request.TopPct = top;
                    }

                    if (!PresentationHostAdapter.TryManageShape(
                            dest,
                            source,
                            request,
                            out PresentationManageShapeResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["slide_id"] = hostResult.SlideId ?? "",
                        ["warnings"] = hostResult.Warnings ?? new List<string>(),
                        ["display_contents"] = BuildDisplay(hostResult)
                    };

                    if (string.Equals(hostResult.Action, "delete", StringComparison.Ordinal))
                    {
                        data["deleted_count"] = hostResult.DeletedCount;
                        data["deleted_shape_ids"] = hostResult.DeletedShapeIds ?? new List<string>();
                    }

                    if (!string.IsNullOrEmpty(hostResult.GroupShapeId))
                    {
                        data["group_shape_id"] = hostResult.GroupShapeId;
                    }

                    if (string.Equals(hostResult.Action, "duplicate_group", StringComparison.Ordinal))
                    {
                        data["source_slide_id"] = hostResult.SourceSlideId ?? "";
                        data["source_shape_id"] = hostResult.SourceShapeId ?? "";
                        data["source_channel_id"] = ChannelRegistry.ToPublicId(hostResult.SourceChannelId) ?? "";
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
                    return new ToolResult { Success = false, Error = "管理页内形状失败: " + ex.Message };
                }
            };
        }

        private static bool TryRejectForeignArgs(
            string action,
            Dictionary<string, object> args,
            out string error)
        {
            error = null;
            bool hasSource = HasArgKey(args, "source_channel_id");
            bool hasToSlide = HasArgKey(args, "to_slide_id");
            bool hasLeft = HasArgKey(args, "left");
            bool hasTop = HasArgKey(args, "top");
            bool hasWidth = HasArgKey(args, "width");
            bool hasHeight = HasArgKey(args, "height");
            bool hasShapeId = HasArgKey(args, "shape_id");
            bool hasShapeIds = HasArgKey(args, "shape_ids");

            switch (action)
            {
                case "delete":
                    if (hasSource || hasToSlide || hasLeft || hasTop || hasWidth || hasHeight)
                    {
                        error = "delete 不接受 source_channel_id / to_slide_id / left / top / width / height";
                        return false;
                    }

                    return true;
                case "group":
                    if (hasSource || hasToSlide || hasShapeId || hasLeft || hasTop || hasWidth || hasHeight)
                    {
                        error = "group 不接受 source_channel_id / to_slide_id / shape_id / left / top / width / height";
                        return false;
                    }

                    return true;
                case "duplicate_group":
                    if (hasShapeIds || hasWidth || hasHeight)
                    {
                        error = "duplicate_group 不接受 shape_ids / width / height";
                        return false;
                    }

                    return true;
                default:
                    error = "action 须为 delete、group 或 duplicate_group";
                    return false;
            }
        }

        private static bool TryResolveSourceChannel(
            string action,
            IOperationChannel dest,
            Dictionary<string, object> args,
            out IOperationChannel source,
            out string error)
        {
            source = dest;
            error = null;
            if (!string.Equals(action, "duplicate_group", StringComparison.Ordinal))
            {
                return true;
            }

            if (!HasArgKey(args, "source_channel_id"))
            {
                return true;
            }

            string raw = GetStringArg(args, "source_channel_id");
            if (string.IsNullOrEmpty(raw))
            {
                error = "未知 source_channel_id: ";
                return false;
            }

            if (!ChannelRegistry.TryGet(raw, out source) || source == null)
            {
                error = "未知 source_channel_id: " + raw;
                return false;
            }

            return true;
        }

        private static bool TryParsePct(
            Dictionary<string, object> args,
            string key,
            out double value,
            out string error)
        {
            value = 0;
            error = null;
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                error = key + " 须为相对目标幻灯片的百分比";
                return false;
            }

            string raw = Convert.ToString(args[key], CultureInfo.InvariantCulture)?.Trim() ?? "";
            if (raw.EndsWith("%", StringComparison.Ordinal))
            {
                raw = raw.Substring(0, raw.Length - 1).Trim();
            }

            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || double.IsNaN(value)
                || double.IsInfinity(value))
            {
                error = "left/top 须为相对目标幻灯片的百分比";
                return false;
            }

            return true;
        }

        private static string BuildDisplay(PresentationManageShapeResult result)
        {
            var sb = new StringBuilder();
            sb.Append("action=").Append(result?.Action ?? "")
                .Append(" slide_id=").Append(result?.SlideId ?? "");
            if (string.Equals(result?.Action, "delete", StringComparison.Ordinal))
            {
                sb.Append(" deleted=").Append(result?.DeletedCount ?? 0);
            }
            else
            {
                sb.Append(" group_shape_id=").Append(result?.GroupShapeId ?? "");
            }

            if (string.Equals(result?.Action, "duplicate_group", StringComparison.Ordinal))
            {
                sb.Append(" source_slide_id=").Append(result?.SourceSlideId ?? "")
                    .Append(" source_shape_id=").Append(result?.SourceShapeId ?? "");
                sb.Append(" 先 F_read_ppt_html 再改新叶子");
            }

            return sb.ToString();
        }

        private static bool HasArgKey(Dictionary<string, object> args, string key)
        {
            return args != null && args.ContainsKey(key);
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }

        private static string[] TryGetStringArray(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            object raw = args[key];
            if (raw is string[] arr)
            {
                return arr;
            }

            if (raw is IList list)
            {
                var items = new List<string>();
                foreach (object item in list)
                {
                    string s = Convert.ToString(item)?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(s))
                    {
                        items.Add(s);
                    }
                }

                return items.ToArray();
            }

            return null;
        }
    }
}
