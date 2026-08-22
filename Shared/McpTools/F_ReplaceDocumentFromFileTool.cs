using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 从文件替换文档内容（扁平入参：replace_file + original_codes + format）。
    /// 本工具不直触 Word/WPS：转发 <c>channel_id</c> 后委托
    /// <c>F_process_document_actions</c>（已经 DocumentHost），故 WPS 随改字路径可用。
    /// </summary>
    public static class F_ReplaceDocumentFromFileTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_replace_document_from_file"] = async (args) =>
            {
                try
                {
                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "replace_file"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    string originalCodes = args.ContainsKey("original_codes")
                        ? args["original_codes"]?.ToString()?.Trim()
                        : "";

                    if (string.IsNullOrEmpty(originalCodes))
                    {
                        return new ToolResult { Success = false, Error = "必须提供 original_codes 参数（要替换的连续编码序列）" };
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

                    var processArgs = BuildReplaceProcessArgs(loadResult.Text, originalCodes, args["format"]);
                    if (args.ContainsKey("channel_id"))
                    {
                        processArgs["channel_id"] = args["channel_id"];
                    }

                    if (!toolRegistry.TryGetValue("F_process_document_actions", out var processHandler))
                    {
                        return new ToolResult { Success = false, Error = "F_process_document_actions 未注册" };
                    }

                    System.Diagnostics.Debug.WriteLine(
                        "[F_replace_document_from_file] delegate → F_process_document_actions; " +
                        $"channel_id={ChannelContext.TryGetChannelIdFromParameters(processArgs) ?? "(default)"}");

                    ToolResult processResult = await processHandler(processArgs);
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
                            replace_file = resolved.Display,
                            original_codes = originalCodes,
                            characters_replaced = loadResult.Text?.Length ?? 0,
                            process_result = processResult.Data,
                            message = $"已从文件 {resolved.Display} 替换文档内容（编码: {originalCodes}）"
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ERROR] replace_document_from_file 异常: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"从文件替换文档失败: {ex.Message}" };
                }
            };
        }

        private static Dictionary<string, object> BuildReplaceProcessArgs(
            string newContent,
            string originalCodes,
            object format)
        {
            return new Dictionary<string, object>
            {
                ["action"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "replace",
                        ["detail"] = new Dictionary<string, object>
                        {
                            ["original"] = new Dictionary<string, object> { ["codes"] = originalCodes },
                            ["new"] = newContent ?? "",
                            ["format"] = format
                        }
                    }
                }
            };
        }
    }
}
