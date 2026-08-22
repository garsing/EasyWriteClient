using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_PptAnimationTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_ppt_animation"] = async (args) =>
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
                            Error = "须提供 action（list|add|replace_all|clear）"
                        };
                    }

                    var request = new PresentationAnimationRequest
                    {
                        Action = action,
                        SlideId = GetStringArg(args, "slide_id"),
                        ShapeId = GetShapeIdArg(args),
                        Effects = ParseEffects(args, action)
                    };

                    if (string.Equals(action.Trim(), "add", StringComparison.OrdinalIgnoreCase)
                        && (request.Effects == null || request.Effects.Count == 0)
                        && !string.IsNullOrWhiteSpace(request.ShapeId))
                    {
                        request.Effects = new List<PresentationAnimationEffectSpec>
                        {
                            BuildFlatSpec(args, request.ShapeId)
                        };
                    }

                    if (string.Equals(action.Trim(), "add", StringComparison.OrdinalIgnoreCase)
                        && request.Effects != null
                        && request.Effects.Count > 1)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "add 仅支持一条效果；多条请用 replace_all"
                        };
                    }

                    if (!PresentationHostAdapter.TryPptAnimation(
                            channel,
                            request,
                            out PresentationAnimationResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var effects = new List<Dictionary<string, object>>();
                    if (hostResult.Effects != null)
                    {
                        foreach (PresentationAnimationEffectSpec e in hostResult.Effects)
                        {
                            if (e == null)
                            {
                                continue;
                            }

                            effects.Add(SpecToDict(e));
                        }
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["slide_id"] = hostResult.SlideId ?? "",
                        ["effects"] = effects,
                        ["added_count"] = hostResult.AddedCount,
                        ["cleared_count"] = hostResult.ClearedCount,
                        ["display_contents"] = BuildDisplayContents(hostResult)
                    };
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
                    return new ToolResult { Success = false, Error = "页内动画失败: " + ex.Message };
                }
            };
        }

        private static PresentationAnimationEffectSpec BuildFlatSpec(Dictionary<string, object> args, string shapeId)
        {
            return new PresentationAnimationEffectSpec
            {
                ShapeId = shapeId,
                Category = GetStringArg(args, "category"),
                Effect = GetStringArg(args, "effect"),
                Trigger = GetStringArg(args, "trigger"),
                DurationMs = TryGetOptionalInt(args, "duration_ms"),
                DelayMs = TryGetOptionalInt(args, "delay_ms"),
                Dir = GetStringArg(args, "dir"),
                Order = TryGetOptionalInt(args, "order")
            };
        }

        private static List<PresentationAnimationEffectSpec> ParseEffects(Dictionary<string, object> args, string action)
        {
            if (args == null || !args.ContainsKey("effects") || args["effects"] == null)
            {
                if (string.Equals(action?.Trim(), "replace_all", StringComparison.OrdinalIgnoreCase))
                {
                    return new List<PresentationAnimationEffectSpec>();
                }

                return null;
            }

            object raw = args["effects"];
            if (!(raw is IList list))
            {
                return null;
            }

            var result = new List<PresentationAnimationEffectSpec>();
            foreach (object item in list)
            {
                if (item == null)
                {
                    continue;
                }

                if (item is PresentationAnimationEffectSpec typed)
                {
                    result.Add(typed);
                    continue;
                }

                if (item is Dictionary<string, object> dict)
                {
                    result.Add(DictToSpec(dict));
                    continue;
                }

                if (item is IDictionary idict)
                {
                    var mapped = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    foreach (DictionaryEntry entry in idict)
                    {
                        mapped[Convert.ToString(entry.Key) ?? ""] = entry.Value;
                    }

                    result.Add(DictToSpec(mapped));
                }
            }

            return result;
        }

        private static PresentationAnimationEffectSpec DictToSpec(Dictionary<string, object> dict)
        {
            return new PresentationAnimationEffectSpec
            {
                ShapeId = GetFromDict(dict, "ShapeId", "shape_id"),
                Category = GetFromDict(dict, "category"),
                Effect = GetFromDict(dict, "effect"),
                Trigger = GetFromDict(dict, "trigger"),
                DurationMs = TryGetOptionalIntFromDict(dict, "duration_ms"),
                DelayMs = TryGetOptionalIntFromDict(dict, "delay_ms"),
                Dir = GetFromDict(dict, "dir"),
                Order = TryGetOptionalIntFromDict(dict, "order")
            };
        }

        private static Dictionary<string, object> SpecToDict(PresentationAnimationEffectSpec e)
        {
            var d = new Dictionary<string, object>
            {
                ["ShapeId"] = e.ShapeId ?? "",
                ["category"] = e.Category ?? "",
                ["effect"] = e.Effect ?? "",
                ["trigger"] = e.Trigger ?? ""
            };
            if (e.DurationMs.HasValue)
            {
                d["duration_ms"] = e.DurationMs.Value;
            }

            if (e.DelayMs.HasValue)
            {
                d["delay_ms"] = e.DelayMs.Value;
            }

            if (!string.IsNullOrEmpty(e.Dir))
            {
                d["dir"] = e.Dir;
            }

            if (e.Order.HasValue)
            {
                d["order"] = e.Order.Value;
            }

            if (e.EffectIndex.HasValue)
            {
                d["effect_index"] = e.EffectIndex.Value;
            }

            return d;
        }

        private static string BuildDisplayContents(PresentationAnimationResult result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("action=" + (result?.Action ?? ""));
            sb.AppendLine("slide_id=" + (result?.SlideId ?? ""));
            sb.AppendLine("effects=" + (result?.Effects?.Count ?? 0));
            sb.AppendLine();
            sb.AppendLine("| # | ShapeId | category | effect | trigger | dir |");
            sb.AppendLine("|--:|---------|----------|--------|---------|-----|");
            if (result?.Effects != null)
            {
                int i = 1;
                foreach (PresentationAnimationEffectSpec e in result.Effects)
                {
                    if (e == null)
                    {
                        continue;
                    }

                    sb.Append("| ").Append(i++)
                        .Append(" | ").Append(e.ShapeId ?? "")
                        .Append(" | ").Append(e.Category ?? "")
                        .Append(" | ").Append(e.Effect ?? "")
                        .Append(" | ").Append(e.Trigger ?? "")
                        .Append(" | ").Append(e.Dir ?? "")
                        .AppendLine(" |");
                }
            }

            return sb.ToString().TrimEnd();
        }

        private static string GetShapeIdArg(Dictionary<string, object> args)
        {
            string v = GetStringArg(args, "ShapeId");
            if (!string.IsNullOrEmpty(v))
            {
                return v;
            }

            return GetStringArg(args, "shape_id");
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }

        private static string GetFromDict(Dictionary<string, object> dict, params string[] keys)
        {
            if (dict == null)
            {
                return "";
            }

            foreach (string key in keys)
            {
                if (dict.TryGetValue(key, out object raw) && raw != null)
                {
                    return Convert.ToString(raw)?.Trim() ?? "";
                }

                foreach (KeyValuePair<string, object> kv in dict)
                {
                    if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                    {
                        return Convert.ToString(kv.Value)?.Trim() ?? "";
                    }
                }
            }

            return "";
        }

        private static int? TryGetOptionalInt(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            return ParseInt(args[key]);
        }

        private static int? TryGetOptionalIntFromDict(Dictionary<string, object> dict, string key)
        {
            if (dict == null)
            {
                return null;
            }

            if (dict.TryGetValue(key, out object raw) && raw != null)
            {
                return ParseInt(raw);
            }

            foreach (KeyValuePair<string, object> kv in dict)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    return ParseInt(kv.Value);
                }
            }

            return null;
        }

        private static int? ParseInt(object raw)
        {
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
    }
}
