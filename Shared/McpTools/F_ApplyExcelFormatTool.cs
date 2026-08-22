using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ApplyExcelFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_excel_format"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    string sheet = GetStringArg(args, "sheet");
                    if (string.IsNullOrWhiteSpace(sheet))
                    {
                        return new ToolResult { Success = false, Error = "必须提供 sheet" };
                    }

                    string range = GetStringArg(args, "range");
                    if (string.IsNullOrWhiteSpace(range))
                    {
                        return new ToolResult { Success = false, Error = "必须提供 range" };
                    }

                    object formatRaw = null;
                    if (args != null && args.ContainsKey("format"))
                    {
                        formatRaw = args["format"];
                    }

                    if (formatRaw == null)
                    {
                        return new ToolResult { Success = false, Error = "必须提供 format 对象" };
                    }

                    if (!SpreadsheetFormatParse.TryParseFormatObject(
                            CoerceFormatArg(formatRaw),
                            out SpreadsheetFormatPatch patch,
                            out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    var request = new SpreadsheetFormatRequest
                    {
                        SheetName = sheet.Trim(),
                        RangeA1 = range.Trim(),
                        Format = patch
                    };

                    if (!SpreadsheetHostAdapter.TryApplyFormat(
                            channel,
                            request,
                            out SpreadsheetFormatResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                            ["kind"] = hostResult.Kind ?? "",
                            ["sheet"] = hostResult.Sheet ?? "",
                            ["actual_range"] = hostResult.ActualRange ?? "",
                            ["applied_fields"] = hostResult.AppliedFields ?? new List<string>()
                        }
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "套用 Excel 格式失败: " + ex.Message };
                }
            };
        }

        private static object CoerceFormatArg(object raw)
        {
            if (raw == null)
            {
                return null;
            }

            if (raw is Dictionary<string, object>)
            {
                return raw;
            }

            try
            {
                var serializer = new JavaScriptSerializer();
                return serializer.DeserializeObject(Convert.ToString(raw) ?? "");
            }
            catch (Exception)
            {
                return raw;
            }
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }
    }
}
