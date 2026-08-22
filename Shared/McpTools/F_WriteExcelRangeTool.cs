using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_WriteExcelRangeTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_write_excel_range"] = async (args) =>
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

                    bool clear = args != null
                        && args.ContainsKey("clear")
                        && Convert.ToBoolean(args["clear"]);
                    string csvFilename = FilePathResolver.TryGetArg(args, "path", "csv_filename");

                    if (clear && !string.IsNullOrEmpty(csvFilename))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "须提供 path+range，或 clear+range（二者互斥）"
                        };
                    }

                    if (!clear && string.IsNullOrEmpty(csvFilename))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "须提供 path+range，或 clear+range"
                        };
                    }

                    var request = new SpreadsheetWriteRequest
                    {
                        SheetName = sheet.Trim(),
                        RangeA1 = range.Trim(),
                        Mode = clear ? SpreadsheetWriteMode.Clear : SpreadsheetWriteMode.Csv
                    };

                    if (!clear)
                    {
                        if (!FilePathResolver.TryResolve(csvFilename, out ResolvedFilePath csvResolved, out string csvError))
                        {
                            return new ToolResult { Success = false, Error = csvError };
                        }

                        var ensure = await FilePathResolver.ReadBytesAsync(csvResolved).ConfigureAwait(false);
                        if (!ensure.Success)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = ensure.Error ?? ("CSV 不存在: " + csvFilename)
                            };
                        }

                        List<List<string>> parsed = SpreadsheetCsv.ParseFile(csvResolved.LocalPath);
                        if (parsed == null)
                        {
                            return new ToolResult { Success = false, Error = "CSV 文件解析失败" };
                        }

                        request.CsvGrid = parsed;
                    }

                    if (!SpreadsheetHostAdapter.TryWriteRange(
                            channel,
                            request,
                            out SpreadsheetWriteResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = hostResult.ChannelId ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["sheet"] = hostResult.Sheet ?? "",
                        ["mode"] = hostResult.Mode ?? "",
                        ["written_count"] = hostResult.WrittenCount,
                        ["actual_range"] = hostResult.ActualRange ?? ""
                    };
                    if (!clear)
                    {
                        data["csv_filename"] = csvFilename;
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
                    return new ToolResult { Success = false, Error = "写入 Excel 区域失败: " + ex.Message };
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
