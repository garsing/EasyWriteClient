using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ReadExcelRangeTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_read_excel_range"] = async (args) =>
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
                    bool includeFormulas = args != null
                        && args.ContainsKey("include_formulas")
                        && Convert.ToBoolean(args["include_formulas"]);

                    if (!SpreadsheetHostAdapter.TryReadRange(
                            channel,
                            sheet.Trim(),
                            range,
                            includeFormulas,
                            out SpreadsheetRangeResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    string table = WorkbookPreviewMarkup.BuildTable(
                        hostResult.Sheet,
                        hostResult.ActualRange ?? "",
                        hostResult.Truncated,
                        hostResult.Rows);

                    string display = "sheet=" + (hostResult.Sheet ?? "")
                        + " actual_range=" + (hostResult.ActualRange ?? "")
                        + "\n"
                        + table;

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
                            ["truncated"] = hostResult.Truncated,
                            ["truncated_reason"] = hostResult.TruncatedReason ?? "",
                            ["display_contents"] = display,
                            ["table"] = table,
                            ["include_formulas"] = hostResult.IncludeFormulas
                        }
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "读取 Excel 区域失败: " + ex.Message };
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
