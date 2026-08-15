using System;
using System.Collections.Generic;
using System.IO;
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
                    string contentRaw = GetStringArg(args, "content");
                    if (!SpreadsheetContentModeUtil.TryParse(contentRaw, out SpreadsheetContentMode contentMode, out string contentError))
                    {
                        return new ToolResult { Success = false, Error = contentError };
                    }

                    string exportCsv = GetStringArg(args, "export_csv");
                    if (!string.IsNullOrEmpty(exportCsv))
                    {
                        exportCsv = Path.GetFileName(exportCsv.Trim());
                        if (string.IsNullOrEmpty(exportCsv)
                            || exportCsv.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                        {
                            return new ToolResult { Success = false, Error = "export_csv 须为合法裸文件名" };
                        }
                    }

                    if (!SpreadsheetHostAdapter.TryReadRange(
                            channel,
                            sheet.Trim(),
                            range,
                            contentMode,
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
                        + " content=" + SpreadsheetContentModeUtil.ToWire(contentMode)
                        + "\n"
                        + table;

                    var data = new Dictionary<string, object>
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
                        ["content"] = SpreadsheetContentModeUtil.ToWire(contentMode)
                    };

                    if (!string.IsNullOrEmpty(exportCsv))
                    {
                        List<List<string>> grid = SpreadsheetCsv.BuildGridFromPreview(
                            hostResult.ActualRange ?? "",
                            hostResult.Rows,
                            out int csvRows,
                            out int csvCols,
                            out string gridError);
                        if (grid == null)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = "导出 CSV 失败: " + (gridError ?? "无法构建网格")
                            };
                        }

                        string localPath = WorkspacePathResolver.ResolveWritePath(exportCsv);
                        try
                        {
                            SpreadsheetCsv.WriteFile(localPath, grid);
                        }
                        catch (Exception ex)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = "导出 CSV 失败: " + ex.Message
                            };
                        }

                        bool uploaded = await McpToolsHelpers.UploadWorkspaceFileAsync(localPath, exportCsv)
                            .ConfigureAwait(false);
                        if (!uploaded)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = "导出 CSV 已写本地但上传工作区失败: " + exportCsv
                            };
                        }

                        data["csv_filename"] = exportCsv;
                        data["csv_rows"] = csvRows;
                        data["csv_columns"] = csvCols;
                    }

                    return new ToolResult { Success = true, Data = data };
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
