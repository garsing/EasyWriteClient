using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ManageExcelSheetTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_manage_excel_sheet"] = async (args) =>
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
                        return new ToolResult { Success = false, Error = "须提供 action（add|rename|delete）" };
                    }

                    bool confirm = args != null
                        && args.ContainsKey("confirm")
                        && Convert.ToBoolean(args["confirm"]);

                    var request = new SpreadsheetManageSheetRequest
                    {
                        Action = action,
                        Name = GetStringArg(args, "name"),
                        Sheet = GetStringArg(args, "sheet"),
                        Before = GetStringArg(args, "before"),
                        Confirm = confirm
                    };

                    if (!SpreadsheetHostAdapter.TryManageSheet(
                            channel,
                            request,
                            out SpreadsheetManageSheetResult hostResult,
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
                            ["channel_id"] = hostResult.ChannelId ?? "",
                            ["kind"] = hostResult.Kind ?? "",
                            ["action"] = hostResult.Action ?? "",
                            ["sheet"] = hostResult.Sheet ?? "",
                            ["sheets"] = hostResult.Sheets ?? new List<string>()
                        }
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "管理工作表失败: " + ex.Message };
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
