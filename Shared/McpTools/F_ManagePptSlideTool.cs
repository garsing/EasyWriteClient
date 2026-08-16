using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_ManagePptSlideTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_manage_ppt_slide"] = async (args) =>
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
                        return new ToolResult
                        {
                            Success = false,
                            Error = "须提供 action（add|duplicate|delete|move|reorder）"
                        };
                    }

                    bool confirm = args != null
                        && args.ContainsKey("confirm")
                        && Convert.ToBoolean(args["confirm"]);

                    var request = new PresentationManageSlideRequest
                    {
                        Action = action,
                        SlideId = GetStringArg(args, "slide_id"),
                        Layout = GetStringArg(args, "layout"),
                        Confirm = confirm,
                        ToIndex = TryGetOptionalInt(args, "to_index"),
                        Order = TryGetStringArray(args, "order")
                    };

                    if (!PresentationHostAdapter.TryManageSlide(
                            channel,
                            request,
                            out PresentationManageSlideResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var slides = new List<Dictionary<string, object>>();
                    if (hostResult.Slides != null)
                    {
                        foreach (PresentationSlideInfo slide in hostResult.Slides)
                        {
                            if (slide == null)
                            {
                                continue;
                            }

                            slides.Add(new Dictionary<string, object>
                            {
                                ["index"] = slide.Index,
                                ["slide_id"] = slide.SlideId ?? "",
                                ["title"] = slide.Title ?? "",
                                ["layout"] = slide.Layout ?? "",
                                ["hidden"] = slide.Hidden
                            });
                        }
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = hostResult.ChannelId ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["slide_id"] = hostResult.FocusSlideId ?? "",
                        ["slide_count"] = hostResult.SlideCount,
                        ["slides"] = slides,
                        ["display_contents"] = BuildDisplayContents(hostResult)
                    };
                    if (hostResult.FocusIndex.HasValue)
                    {
                        data["to_index"] = hostResult.FocusIndex.Value;
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = data
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "管理幻灯片失败: " + ex.Message };
                }
            };
        }

        private static string BuildDisplayContents(PresentationManageSlideResult result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("action=" + (result?.Action ?? ""));
            if (!string.IsNullOrEmpty(result?.FocusSlideId))
            {
                sb.Append("slide_id=").Append(result.FocusSlideId);
                if (result.FocusIndex.HasValue)
                {
                    sb.Append(" to_index=").Append(result.FocusIndex.Value);
                }

                sb.AppendLine();
            }

            sb.AppendLine("slide_count=" + (result?.SlideCount ?? 0));
            sb.AppendLine();
            sb.AppendLine("| index | slide_id | layout | hidden |");
            sb.AppendLine("|------:|---------:|--------|--------|");
            if (result?.Slides != null)
            {
                foreach (PresentationSlideInfo slide in result.Slides)
                {
                    if (slide == null)
                    {
                        continue;
                    }

                    sb.Append("| ").Append(slide.Index)
                        .Append(" | ").Append(slide.SlideId ?? "")
                        .Append(" | ").Append(slide.Layout ?? "")
                        .Append(" | ").Append(slide.Hidden ? "true" : "false")
                        .AppendLine(" |");
                }
            }

            return sb.ToString().TrimEnd();
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }

        private static int? TryGetOptionalInt(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            object raw = args[key];
            if (raw is int i)
            {
                return i;
            }

            if (raw is long l)
            {
                return (int)l;
            }

            if (raw is double d)
            {
                return (int)d;
            }

            if (int.TryParse(Convert.ToString(raw), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
            {
                return parsed;
            }

            return null;
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
                    items.Add(Convert.ToString(item)?.Trim() ?? "");
                }

                return items.ToArray();
            }

            return null;
        }
    }
}
