using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_GetWorkbookContentTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_workbook_content"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!SpreadsheetHostAdapter.TryGetWorkbookContent(
                            channel,
                            out WorkbookContentResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var sheets = new List<Dictionary<string, object>>();
                    if (hostResult.Sheets != null)
                    {
                        foreach (WorkbookSheetInfo sheet in hostResult.Sheets)
                        {
                            if (sheet == null)
                            {
                                continue;
                            }

                            sheets.Add(new Dictionary<string, object>
                            {
                                ["name"] = sheet.Name ?? "",
                                ["hidden"] = sheet.Hidden,
                                ["sheet_type"] = sheet.SheetType ?? "worksheet",
                                ["used_range"] = sheet.UsedRange ?? "",
                                ["last_row"] = sheet.LastRow,
                                ["last_col"] = sheet.LastCol ?? "",
                                ["preview"] = sheet.Preview ?? "",
                                ["preview_truncated"] = sheet.PreviewTruncated
                            });
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
                            ["name"] = hostResult.Name ?? "",
                            ["path"] = hostResult.Path ?? "",
                            ["display_contents"] = WorkbookPreviewMarkup.BuildDisplayContents(hostResult),
                            ["sheets"] = sheets
                        }
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "读取工作簿概览失败: " + ex.Message };
                }
            };
        }
    }
}
