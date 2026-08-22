using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 从文件插入文档内容（扁平入参：insert_file + format + 可选定位）。
    /// 本工具不直触 Word/WPS：转发 <c>channel_id</c> 后委托
    /// <c>F_process_document_actions</c>（已经 DocumentHost），故 WPS 随改字路径可用。
    /// </summary>
    public static class F_InsertDocumentFromFileTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_insert_document_from_file"] = async (args) =>
            {
                try
                {
                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "insert_file"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    ActionFormatHelper.Spec formatSpec = ActionFormatHelper.ParseFormatObject(args, required: true);
                    if (!formatSpec.Ok)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = formatSpec.Error ?? "必须指定 format（inherit_from 或 char_format）"
                        };
                    }

                    var loadResult = await FilePathResolver.ReadAsync(resolved).ConfigureAwait(false);
                    if (!loadResult.Success || string.IsNullOrEmpty(loadResult.Text))
                    {
                        return new ToolResult { Success = false, Error = loadResult.Error ?? $"文件内容为空: {resolved.Display}" };
                    }

                    string targetCodes = args.ContainsKey("target_codes") ? args["target_codes"]?.ToString()?.Trim() : "";
                    string inTable = args.ContainsKey("in_table") ? args["in_table"]?.ToString()?.Trim() : "";
                    string targetTable = args.ContainsKey("target_table") ? args["target_table"]?.ToString()?.Trim() : "";
                    string targetChart = args.ContainsKey("target_chart") ? args["target_chart"]?.ToString()?.Trim() : "";
                    string targetImage = args.ContainsKey("target_image") ? args["target_image"]?.ToString()?.Trim() : "";

                    var build = BuildInsertProcessArgs(
                        loadResult.Text,
                        targetCodes,
                        inTable,
                        targetTable,
                        targetChart,
                        targetImage,
                        args["format"]);
                    if (!string.IsNullOrEmpty(build.error))
                    {
                        return new ToolResult { Success = false, Error = build.error };
                    }

                    if (!toolRegistry.TryGetValue("F_process_document_actions", out var processHandler))
                    {
                        return new ToolResult { Success = false, Error = "F_process_document_actions 未注册" };
                    }

                    if (args.ContainsKey("channel_id"))
                    {
                        build.args["channel_id"] = args["channel_id"];
                    }

                    System.Diagnostics.Debug.WriteLine(
                        "[F_insert_document_from_file] delegate → F_process_document_actions; " +
                        $"channel_id={ChannelContext.TryGetChannelIdFromParameters(build.args) ?? "(default)"}");

                    ToolResult processResult = await processHandler(build.args);
                    if (!processResult.Success)
                    {
                        return processResult;
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            path = resolved.Display,
                            insert_file = resolved.Display,
                            target_codes = string.IsNullOrEmpty(targetCodes) ? null : targetCodes,
                            in_table = string.IsNullOrEmpty(inTable) ? null : inTable,
                            target_table = string.IsNullOrEmpty(targetTable) ? null : targetTable,
                            target_chart = string.IsNullOrEmpty(targetChart) ? null : targetChart,
                            target_image = string.IsNullOrEmpty(targetImage) ? null : targetImage,
                            characters_inserted = loadResult.Text?.Length ?? 0,
                            process_result = processResult.Data,
                            message = $"已从文件 {resolved.Display} 插入文档内容"
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ERROR] insert_document_from_file 异常: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"从文件插入文档失败: {ex.Message}" };
                }
            };
        }

        private static (Dictionary<string, object> args, string error) BuildInsertProcessArgs(
            string insertContent,
            string targetCodes,
            string inTable,
            string targetTable,
            string targetChart,
            string targetImage,
            object format)
        {
            var detail = new Dictionary<string, object>
            {
                ["insert"] = insertContent ?? "",
                ["format"] = format
            };

            bool hasCodes = !string.IsNullOrEmpty(targetCodes);
            bool hasTable = !string.IsNullOrEmpty(targetTable);
            bool hasChart = !string.IsNullOrEmpty(targetChart);
            bool hasImage = !string.IsNullOrEmpty(targetImage);
            int modeCount = (hasCodes ? 1 : 0) + (hasTable ? 1 : 0) + (hasChart ? 1 : 0) + (hasImage ? 1 : 0);
            if (modeCount > 1)
            {
                return (null, "target_codes / target_table / target_chart / target_image 四选一");
            }

            if (!string.IsNullOrEmpty(inTable) && !hasCodes)
            {
                return (null, "in_table 仅可与 target_codes 同用");
            }

            if (hasCodes)
            {
                var target = new Dictionary<string, object> { ["codes"] = targetCodes };
                if (!string.IsNullOrEmpty(inTable))
                {
                    target["in_table"] = inTable;
                }

                detail["target"] = target;
            }
            else if (hasTable)
            {
                detail["target"] = new Dictionary<string, object> { ["table"] = targetTable };
            }
            else if (hasChart)
            {
                detail["target"] = new Dictionary<string, object> { ["chart"] = targetChart };
            }
            else if (hasImage)
            {
                detail["target"] = new Dictionary<string, object> { ["image"] = targetImage };
            }

            return (new Dictionary<string, object>
            {
                ["action"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "insert",
                        ["detail"] = detail
                    }
                }
            }, null);
        }
    }
}
