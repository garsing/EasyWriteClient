using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_PptTransitionTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_ppt_transition"] = async (args) =>
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
                            Error = "须提供 action（list|get|set|clear）"
                        };
                    }

                    string actionNorm = action.Trim().ToLowerInvariant();
                    bool listHasForbiddenFilter = HasArg(args, "slide_ids")
                        || (string.Equals(actionNorm, "list", StringComparison.Ordinal)
                            && HasArg(args, "slide_id"));

                    var request = new PresentationTransitionRequest
                    {
                        Action = action,
                        SlideId = GetStringArg(args, "slide_id"),
                        Effect = GetStringArg(args, "effect"),
                        Dir = GetStringArg(args, "dir"),
                        DurationMs = TryGetOptionalInt(args, "duration_ms"),
                        AdvanceOnClick = TryGetOptionalBool(args, "advance_on_click"),
                        AdvanceAfterMs = TryGetOptionalInt(args, "advance_after_ms"),
                        ListHasForbiddenFilter = listHasForbiddenFilter
                    };

                    if (!PresentationHostAdapter.TryPptTransition(
                            channel,
                            request,
                            out PresentationTransitionResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var transitions = new List<Dictionary<string, object>>();
                    if (hostResult.Transitions != null)
                    {
                        foreach (PresentationTransitionInfo t in hostResult.Transitions)
                        {
                            if (t == null)
                            {
                                continue;
                            }

                            transitions.Add(InfoToDict(t));
                        }
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["transitions"] = transitions,
                        ["display_contents"] = BuildDisplayContents(hostResult)
                    };
                    if (!string.IsNullOrEmpty(hostResult.SlideId))
                    {
                        data["slide_id"] = hostResult.SlideId;
                    }

                    if (hostResult.Warnings != null && hostResult.Warnings.Count > 0)
                    {
                        data["warnings"] = hostResult.Warnings;
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
                    return new ToolResult { Success = false, Error = "换页过渡失败: " + ex.Message };
                }
            };
        }

        private static Dictionary<string, object> InfoToDict(PresentationTransitionInfo t)
        {
            var d = new Dictionary<string, object>
            {
                ["slide_id"] = t.SlideId ?? "",
                ["index"] = t.Index,
                ["effect"] = t.Effect ?? "none",
                ["advance_on_click"] = t.AdvanceOnClick
            };
            if (!string.IsNullOrEmpty(t.Dir))
            {
                d["dir"] = t.Dir;
            }

            if (t.DurationMs.HasValue)
            {
                d["duration_ms"] = t.DurationMs.Value;
            }

            if (t.AdvanceAfterMs.HasValue)
            {
                d["advance_after_ms"] = t.AdvanceAfterMs.Value;
            }

            return d;
        }

        private static string BuildDisplayContents(PresentationTransitionResult result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("| slide_id | index | effect | dir | duration_ms | advance_on_click | advance_after_ms |");
            sb.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");
            if (result?.Transitions != null)
            {
                foreach (PresentationTransitionInfo t in result.Transitions)
                {
                    if (t == null)
                    {
                        continue;
                    }

                    sb.Append("| ")
                        .Append(t.SlideId ?? "")
                        .Append(" | ")
                        .Append(t.Index)
                        .Append(" | ")
                        .Append(t.Effect ?? "")
                        .Append(" | ")
                        .Append(t.Dir ?? "")
                        .Append(" | ")
                        .Append(t.DurationMs.HasValue ? t.DurationMs.Value.ToString(CultureInfo.InvariantCulture) : "")
                        .Append(" | ")
                        .Append(t.AdvanceOnClick ? "true" : "false")
                        .Append(" | ")
                        .Append(t.AdvanceAfterMs.HasValue
                            ? t.AdvanceAfterMs.Value.ToString(CultureInfo.InvariantCulture)
                            : "")
                        .AppendLine(" |");
                }
            }

            sb.AppendLine();
            sb.AppendLine("说明：换页过渡挂靠 slide_id；list=全表，get=单页；不管页内动画（F_ppt_animation）。");
            return sb.ToString();
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

        private static int? TryGetOptionalInt(Dictionary<string, object> args, string key)
        {
            if (!HasArg(args, key))
            {
                return null;
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
                return null;
            }

            if (raw is int i)
            {
                return i;
            }

            if (raw is long l)
            {
                return (int)l;
            }

            if (int.TryParse(Convert.ToString(raw), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
            {
                return parsed;
            }

            return null;
        }

        private static bool? TryGetOptionalBool(Dictionary<string, object> args, string key)
        {
            if (!HasArg(args, key))
            {
                return null;
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
                return null;
            }

            if (raw is bool b)
            {
                return b;
            }

            string s = Convert.ToString(raw)?.Trim();
            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase)
                || s == "1")
            {
                return true;
            }

            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)
                || s == "0")
            {
                return false;
            }

            return null;
        }
    }
}
