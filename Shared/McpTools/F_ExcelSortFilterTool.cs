using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ExcelSortFilterTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_excel_sort_filter"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!SpreadsheetSortFilterParse.TryParseRequest(
                            args,
                            out SpreadsheetSortFilterRequest request,
                            out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    if (!SpreadsheetHostAdapter.TrySortFilter(
                            channel,
                            request,
                            out SpreadsheetSortFilterResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["sheet"] = hostResult.Sheet ?? request.Sheet ?? ""
                    };

                    if (hostResult.Action == "list")
                    {
                        var panels = new List<Dictionary<string, object>>();
                        if (hostResult.Panels != null)
                        {
                            foreach (SortFilterPanel panel in hostResult.Panels)
                            {
                                panels.Add(SpreadsheetSortFilterParse.PanelToWire(panel));
                            }
                        }

                        data["panels"] = panels;
                    }
                    else if (hostResult.Action == "clear")
                    {
                        data["index"] = hostResult.ClearedIndex;
                        data["range"] = hostResult.ClearedRange ?? "";
                    }
                    else if (hostResult.Panel != null)
                    {
                        foreach (KeyValuePair<string, object> pair in SpreadsheetSortFilterParse.PanelToWire(hostResult.Panel))
                        {
                            data[pair.Key] = pair.Value;
                        }
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
                        Error = "排序和筛选失败: " + ex.Message
                    };
                }
            };
        }

        public static bool TryGetContractParseFailure(Dictionary<string, object> args, out ToolResult failure)
        {
            failure = null;
            if (!SpreadsheetSortFilterParse.TryParseRequest(args, out _, out string parseError))
            {
                failure = new ToolResult { Success = false, Error = parseError };
                return true;
            }

            return false;
        }
    }
}
