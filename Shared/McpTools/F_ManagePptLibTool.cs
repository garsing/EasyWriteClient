using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_ManagePptLibTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_manage_ppt_lib"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    string op = GetStringArg(args, "op");
                    if (string.IsNullOrWhiteSpace(op))
                    {
                        op = GetStringArg(args, "action");
                    }

                    ToolResult result;
                    switch ((op ?? "").Trim())
                    {
                        case "listUser":
                            result = ListUser();
                            break;
                        case "getThumbs":
                            result = GetThumbs(args);
                            break;
                        case "writeUser":
                            result = WriteUser(args);
                            break;
                        case "extract":
                            result = Extract(args);
                            break;
                        case "deleteUser":
                            result = DeleteUser(args);
                            break;
                        default:
                            result = new ToolResult
                            {
                                Success = false,
                                Error = "未知 op: " + (op ?? "")
                            };
                            break;
                    }

                    await Task.CompletedTask;
                    return result;
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "组件库操作失败: " + ex.Message };
                }
            };
        }

        private static ToolResult ListUser()
        {
            if (!PptLibStore.TryGetUserRoot(out string root, out string error))
            {
                return new ToolResult { Success = false, Error = error };
            }

            return new ToolResult
            {
                Success = true,
                Data = new Dictionary<string, object>
                {
                    ["items"] = PptLibStore.ScanUser(root)
                }
            };
        }

        private static ToolResult GetThumbs(Dictionary<string, object> args)
        {
            string[] ids = TryGetStringArray(args, "ids");
            if (!PptLibStore.TryGetThumbs(ids, out List<Dictionary<string, object>> items, out string error))
            {
                return new ToolResult { Success = false, Error = error };
            }

            return new ToolResult
            {
                Success = true,
                Data = new Dictionary<string, object> { ["items"] = items }
            };
        }

        private static ToolResult WriteUser(Dictionary<string, object> args)
        {
            string id = GetStringArg(args, "id");
            string category = GetStringArg(args, "category");
            JObject meta = ToJObject(args == null ? null : args.ContainsKey("meta") ? args["meta"] : null);
            string html = GetStringArg(args, "component_html");
            byte[] thumb = TryDecodeBase64(GetStringArg(args, "thumb_base64"));
            bool overwrite = GetBoolArg(args, "overwrite", false);
            if (!PptLibStore.TryWriteUser(id, category, meta, html, thumb, overwrite, out bool overwritten, out string error))
            {
                return new ToolResult { Success = false, Error = error };
            }

            return new ToolResult
            {
                Success = true,
                Data = new Dictionary<string, object>
                {
                    ["id"] = id,
                    ["category"] = category,
                    ["overwritten"] = overwritten,
                    ["has_thumb"] = thumb != null && thumb.Length > 0
                }
            };
        }

        private static ToolResult Extract(Dictionary<string, object> args)
        {
            if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
            {
                return new ToolResult { Success = false, Error = resolveError };
            }

            string shapeId = GetStringArg(args, "shape_id");
            if (string.IsNullOrWhiteSpace(shapeId))
            {
                string[] ids = TryGetStringArray(args, "shape_ids");
                if (ids != null && ids.Length == 1)
                {
                    shapeId = ids[0];
                }
            }

            if (!PptLibExtract.TryExtract(
                    channel,
                    GetStringArg(args, "slide_id"),
                    shapeId,
                    GetStringArg(args, "name"),
                    GetStringArg(args, "description"),
                    GetStringArg(args, "category"),
                    GetStringArg(args, "id"),
                    GetBoolArg(args, "overwrite", false),
                    out Dictionary<string, object> data,
                    out ToolResult errorResult))
            {
                return errorResult;
            }

            return new ToolResult { Success = true, Data = data };
        }

        private static ToolResult DeleteUser(Dictionary<string, object> args)
        {
            string[] ids = TryGetStringArray(args, "ids");
            if (ids == null || ids.Length == 0)
            {
                string single = GetStringArg(args, "id");
                if (!string.IsNullOrWhiteSpace(single))
                {
                    ids = new[] { single };
                }
            }

            if (!PptLibStore.TryDeleteUser(ids, out List<string> deleted, out string error))
            {
                return new ToolResult { Success = false, Error = error };
            }

            return new ToolResult
            {
                Success = true,
                Data = new Dictionary<string, object>
                {
                    ["deleted_count"] = deleted.Count,
                    ["deleted_ids"] = deleted
                }
            };
        }

        private static JObject ToJObject(object raw)
        {
            if (raw == null)
            {
                return null;
            }

            if (raw is JObject jo)
            {
                return jo;
            }

            if (raw is string s)
            {
                return string.IsNullOrWhiteSpace(s) ? null : JObject.Parse(s);
            }

            if (raw is IDictionary dict)
            {
                return JObject.FromObject(dict);
            }

            return JObject.FromObject(raw);
        }

        private static byte[] TryDecodeBase64(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            try
            {
                return Convert.FromBase64String(raw.Trim());
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool GetBoolArg(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return defaultValue;
            }

            object raw = args[key];
            if (raw is bool b)
            {
                return b;
            }

            string s = Convert.ToString(raw)?.Trim();
            if (string.Equals(s, "1", StringComparison.Ordinal) || string.Equals(s, "true", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(s, "0", StringComparison.Ordinal) || string.Equals(s, "false", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return defaultValue;
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
