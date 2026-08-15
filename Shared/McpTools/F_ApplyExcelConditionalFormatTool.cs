using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ApplyExcelConditionalFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_excel_conditional_format"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!SpreadsheetConditionalFormatParse.TryParseRequest(
                            args,
                            out SpreadsheetConditionalFormatRequest request,
                            out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    if (!SpreadsheetHostAdapter.TryApplyConditionalFormat(
                            channel,
                            request,
                            out SpreadsheetConditionalFormatResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = hostResult.ChannelId ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["sheet"] = hostResult.Sheet ?? "",
                        ["requested_range"] = hostResult.RequestedRange ?? "",
                        ["actual_range"] = hostResult.ActualRange ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["rules_count"] = hostResult.RulesCount
                    };

                    if (hostResult.Action == "list"
                        || hostResult.Action == "clear"
                        || hostResult.Action == "add"
                        || hostResult.Action == "replace_all")
                    {
                        var rules = new List<Dictionary<string, object>>();
                        if (hostResult.Rules != null)
                        {
                            foreach (SpreadsheetConditionalRuleInfo info in hostResult.Rules)
                            {
                                rules.Add(SpreadsheetConditionalFormatParse.RuleInfoToWire(info));
                            }
                        }

                        data["rules"] = rules;
                    }

                    if (hostResult.Action == "clear" || hostResult.Action == "replace_all")
                    {
                        data["cleared_count"] = hostResult.ClearedCount;
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
                    return new ToolResult
                    {
                        Success = false,
                        Error = "条件格式操作失败: " + ex.Message
                    };
                }
            };
        }
    }
}
