using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ExcelChartTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_excel_chart"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!SpreadsheetChartParse.TryParseRequest(
                            args,
                            out SpreadsheetChartRequest request,
                            out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    if (!SpreadsheetHostAdapter.TryChart(
                            channel,
                            request,
                            out SpreadsheetChartResult hostResult,
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
                        var charts = new List<Dictionary<string, object>>();
                        if (hostResult.Charts != null)
                        {
                            foreach (SpreadsheetChartInfo info in hostResult.Charts)
                            {
                                charts.Add(SpreadsheetChartParse.ChartInfoToWire(info));
                            }
                        }

                        data["charts"] = charts;
                        data["charts_count"] = hostResult.ChartsCount;
                    }
                    else if (hostResult.Action == "extract")
                    {
                        data["chart_name"] = hostResult.ChartName ?? "";
                        data["sheet"] = hostResult.Sheet ?? "";
                        data["scope"] = hostResult.Scope ?? "";
                        data["xml_content"] = hostResult.XmlContent ?? "";
                        data["dest_range"] = hostResult.DestRange ?? "";
                        if (!string.IsNullOrEmpty(hostResult.Type))
                        {
                            data["type"] = hostResult.Type;
                        }
                    }
                    else
                    {
                        data["chart_name"] = hostResult.ChartName ?? "";
                        data["sheet"] = hostResult.Sheet ?? "";
                        data["dest_range"] = hostResult.DestRange ?? "";
                        if (!string.IsNullOrEmpty(hostResult.Type))
                        {
                            data["type"] = hostResult.Type;
                        }
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
                    return new ToolResult
                    {
                        Success = false,
                        Error = "图表操作失败: " + ex.Message
                    };
                }
            };
        }
    }
}
