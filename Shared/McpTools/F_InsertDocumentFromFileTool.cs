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
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

                    FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out _, "path", "insert_file");
                    DocumentCheckpointMutatingTools.TryReadLocalText(resolved, out string fileText, out _);

                    string targetCodes = args.ContainsKey("target_codes") ? args["target_codes"]?.ToString()?.Trim() : "";
                    string inTable = args.ContainsKey("in_table") ? args["in_table"]?.ToString()?.Trim() : "";
                    string targetTable = args.ContainsKey("target_table") ? args["target_table"]?.ToString()?.Trim() : "";
                    string targetChart = args.ContainsKey("target_chart") ? args["target_chart"]?.ToString()?.Trim() : "";
                    string targetImage = args.ContainsKey("target_image") ? args["target_image"]?.ToString()?.Trim() : "";
                    string targetPosition = args.ContainsKey("target_position") ? args["target_position"]?.ToString()?.Trim() : "";

                    var build = BuildInsertProcessArgs(
                        fileText,
                        targetCodes,
                        inTable,
                        targetTable,
                        targetChart,
                        targetImage,
                        targetPosition,
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
                            target_position = string.IsNullOrEmpty(targetPosition) ? null : targetPosition,
                            characters_inserted = fileText?.Length ?? 0,
                            process_result = processResult.Data,
                            message = InsertTargetLocatorHelper.IsStartPosition(targetPosition)
                                ? $"已从文件 {resolved.Display} 插在正文最前"
                                : $"已从文件 {resolved.Display} 插入文档内容"
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

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "insert_file"))
            {
                failure = DocumentCheckpointMutatingTools.Fail(pathError);
                return true;
            }

            ActionFormatHelper.Spec formatSpec = ActionFormatHelper.ParseFormatObject(args, required: true);
            if (!formatSpec.Ok)
            {
                failure = DocumentCheckpointMutatingTools.Fail(
                    formatSpec.Error ?? ActionFormatHelper.MissingFormatError);
                return true;
            }

            if (!DocumentCheckpointMutatingTools.TryReadLocalText(resolved, out string fileText, out string readError))
            {
                failure = DocumentCheckpointMutatingTools.Fail(readError);
                return true;
            }

            string targetCodes = args.ContainsKey("target_codes") ? args["target_codes"]?.ToString()?.Trim() : "";
            string inTable = args.ContainsKey("in_table") ? args["in_table"]?.ToString()?.Trim() : "";
            string targetTable = args.ContainsKey("target_table") ? args["target_table"]?.ToString()?.Trim() : "";
            string targetChart = args.ContainsKey("target_chart") ? args["target_chart"]?.ToString()?.Trim() : "";
            string targetImage = args.ContainsKey("target_image") ? args["target_image"]?.ToString()?.Trim() : "";
            string targetPosition = args.ContainsKey("target_position") ? args["target_position"]?.ToString()?.Trim() : "";
            var build = BuildInsertProcessArgs(
                fileText,
                targetCodes,
                inTable,
                targetTable,
                targetChart,
                targetImage,
                targetPosition,
                args["format"]);
            if (!string.IsNullOrEmpty(build.error))
            {
                failure = DocumentCheckpointMutatingTools.Fail(build.error);
                return true;
            }

            return false;
        }

        private static (Dictionary<string, object> args, string error) BuildInsertProcessArgs(
            string insertContent,
            string targetCodes,
            string inTable,
            string targetTable,
            string targetChart,
            string targetImage,
            string targetPosition,
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
            bool hasPosition = !string.IsNullOrEmpty(targetPosition);
            int modeCount = (hasCodes ? 1 : 0) + (hasTable ? 1 : 0) + (hasChart ? 1 : 0)
                + (hasImage ? 1 : 0) + (hasPosition ? 1 : 0);
            if (modeCount > 1)
            {
                return (null, "target_codes / target_table / target_chart / target_image / target_position 五选一");
            }

            if (!string.IsNullOrEmpty(inTable) && !hasCodes)
            {
                return (null, "in_table 仅可与 target_codes 同用");
            }

            if (hasPosition)
            {
                if (!InsertTargetLocatorHelper.IsStartPosition(targetPosition))
                {
                    return (null, "target_position 仅支持 start");
                }

                detail["target"] = new Dictionary<string, object>
                {
                    ["position"] = InsertTargetLocatorHelper.PositionStart
                };
            }
            else if (hasCodes)
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
