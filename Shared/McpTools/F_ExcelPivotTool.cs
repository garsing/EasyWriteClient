using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ExcelPivotTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_excel_pivot"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!SpreadsheetPivotParse.TryParseRequest(
                            args,
                            out SpreadsheetPivotRequest request,
                            out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    if (!SpreadsheetHostAdapter.TryPivot(
                            channel,
                            request,
                            out SpreadsheetPivotResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? ""
                    };

                    if (hostResult.Action == "list")
                    {
                        var pivots = new List<Dictionary<string, object>>();
                        if (hostResult.Pivots != null)
                        {
                            foreach (SpreadsheetPivotInfo info in hostResult.Pivots)
                            {
                                pivots.Add(SpreadsheetPivotParse.PivotInfoToWire(info));
                            }
                        }

                        data["pivots"] = pivots;
                        data["pivots_count"] = hostResult.PivotsCount;
                    }
                    else
                    {
                        data["name"] = hostResult.Name ?? "";
                        data["sheet"] = hostResult.Sheet ?? "";
                        data["dest_cell"] = hostResult.DestCell ?? "";
                        if (!string.IsNullOrEmpty(hostResult.SourceSheet))
                        {
                            data["source_sheet"] = hostResult.SourceSheet;
                        }

                        if (!string.IsNullOrEmpty(hostResult.SourceRange))
                        {
                            data["source_range"] = hostResult.SourceRange;
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
                        Error = "透视表操作失败: " + ex.Message
                    };
                }
            };
        }
    }
}
