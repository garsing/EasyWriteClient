using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ApplyExcelStructureTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_excel_structure"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!SpreadsheetStructureParse.TryParseRequest(
                            args,
                            out SpreadsheetStructureRequest request,
                            out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    if (!SpreadsheetHostAdapter.TryApplyStructure(
                            channel,
                            request,
                            out SpreadsheetStructureResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? "",
                        ["sheet"] = hostResult.Sheet ?? "",
                        ["range"] = hostResult.Range ?? "",
                        ["affected"] = hostResult.Affected ?? ""
                    };

                    if (hostResult.Count > 0)
                    {
                        data["count"] = hostResult.Count;
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
                        Error = "格子结构操作失败: " + ex.Message
                    };
                }
            };
        }
    }
}
