using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_GetExcelFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_excel_format"] = async (args) =>
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

                    var request = new SpreadsheetGetFormatRequest
                    {
                        SheetName = sheet.Trim(),
                        RangeA1 = range.Trim()
                    };

                    if (!SpreadsheetHostAdapter.TryGetFormat(
                            channel,
                            request,
                            out SpreadsheetGetFormatResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var items = new List<Dictionary<string, object>>();
                    if (hostResult.Items != null)
                    {
                        foreach (SpreadsheetFormatItem item in hostResult.Items)
                        {
                            var wire = new Dictionary<string, object>();
                            if (!string.IsNullOrEmpty(item.Addr))
                            {
                                wire["addr"] = item.Addr;
                            }

                            if (!string.IsNullOrEmpty(item.Col))
                            {
                                wire["col"] = item.Col;
                            }

                            if (!string.IsNullOrEmpty(item.RangeA1))
                            {
                                wire["range"] = item.RangeA1;
                            }

                            if (item.Mixed.HasValue)
                            {
                                wire["mixed"] = item.Mixed.Value;
                                wire["mixed_fields"] = item.MixedFields ?? new List<string>();
                            }

                            wire["format"] = SpreadsheetFormatRead.ToWireDict(item.Format);
                            items.Add(wire);
                        }
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["channel_id"] = hostResult.ChannelId ?? "",
                            ["kind"] = hostResult.Kind ?? "",
                            ["sheet"] = hostResult.Sheet ?? "",
                            ["requested_range"] = hostResult.RequestedRange ?? "",
                            ["actual_range"] = hostResult.ActualRange ?? "",
                            ["mode"] = hostResult.Mode ?? "",
                            ["items"] = items
                        }
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "读取 Excel 格式失败: " + ex.Message };
                }
            };
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
