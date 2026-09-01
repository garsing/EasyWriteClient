using System;
using System.Collections;
using System.Collections.Generic;
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
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    string action = GetStringArg(args, "action");
                    if (string.IsNullOrWhiteSpace(action))
                    {
                        return new ToolResult { Success = false, Error = "须提供 action（delete）" };
                    }

                    string[] shapeIds = TryGetStringArray(args, "shape_ids");
                    if (shapeIds == null || shapeIds.Length == 0)
                    {
                        string single = GetStringArg(args, "shape_id");
                        if (!string.IsNullOrWhiteSpace(single))
                        {
                            shapeIds = new[] { single };
                        }
                    }

                    var request = new PresentationManageShapeRequest
                    {
                        Action = action,
                        SlideId = GetStringArg(args, "slide_id"),
                        ShapeIds = shapeIds
                    };

                    if (!PresentationHostAdapter.TryManageShape(
                            channel,
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
                        ["deleted_count"] = hostResult.DeletedCount,
                        ["deleted_shape_ids"] = hostResult.DeletedShapeIds ?? new List<string>(),
                        ["warnings"] = hostResult.Warnings ?? new List<string>(),
                        ["display_contents"] = BuildDisplay(hostResult)
                    };

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

        private static string BuildDisplay(PresentationManageShapeResult result)
        {
            var sb = new StringBuilder();
            sb.Append("action=").Append(result?.Action ?? "").Append(" slide_id=")
                .Append(result?.SlideId ?? "").Append(" deleted=").Append(result?.DeletedCount ?? 0);
            return sb.ToString();
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
