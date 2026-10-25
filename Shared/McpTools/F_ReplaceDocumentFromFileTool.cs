using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 从文件替换文档内容（定位对象 original 与 process_actions 同构）。
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
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

                    FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out _, "path", "replace_file");
                    DocumentCheckpointMutatingTools.TryReadLocalText(resolved, out string fileText, out _);
                    ActionContentHelper.Payload payload = ActionContentHelper.TryParseFileBody(fileText, args);
                    Dictionary<string, object> original = InsertTargetLocatorHelper.TryGetNamedLocatorDict(args, "original");

                    var processArgs = BuildReplaceProcessArgs(payload, original);
                    if (args.ContainsKey("channel_id"))
                    {
                        processArgs["channel_id"] = args["channel_id"];
                    }

                    if (!toolRegistry.TryGetValue("F_process_document_actions", out var processHandler))
                    {
                        return new ToolResult { Success = false, Error = "F_process_document_actions 未注册" };
                    }

                    EasyWriteDiagnostics.Log(
                        DebugCategory.Tool,
                        "[F_replace_document_from_file] delegate → F_process_document_actions; channel_id="
                        + (ChannelContext.TryGetChannelIdFromParameters(processArgs) ?? "(default)"));

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
                            original,
                            characters_replaced = fileText?.Length ?? 0,
                            process_result = processResult.Data,
                            message = $"已从文件 {resolved.Display} 替换文档内容"
                        }
                    };
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(
                        DebugCategory.Tool,
                        "[ERROR] replace_document_from_file 异常: " + ex.Message);
                    return new ToolResult { Success = false, Error = $"从文件替换文档失败: {ex.Message}" };
                }
            };
        }

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "replace_file"))
            {
                failure = DocumentCheckpointMutatingTools.Fail(pathError);
                return true;
            }

            string retired = InsertTargetLocatorHelper.RetiredFlatLocatorError(
                args, InsertTargetLocatorHelper.FromFileRetiredReplaceKeys);
            if (retired != null)
            {
                failure = DocumentCheckpointMutatingTools.Fail(retired);
                return true;
            }

            Dictionary<string, object> original = InsertTargetLocatorHelper.TryGetNamedLocatorDict(args, "original");
            if (original == null)
            {
                failure = DocumentCheckpointMutatingTools.Fail(
                    "必须提供 original 定位对象（与 process_actions replace 同构）");
                return true;
            }

            var parsed = InsertTargetLocatorHelper.ParseLocatorObject(
                original, sentenceOnly: true, fieldLabel: "original");
            if (!parsed.Success)
            {
                failure = DocumentCheckpointMutatingTools.Fail(parsed.Error);
                return true;
            }

            if (!DocumentCheckpointMutatingTools.TryReadLocalText(resolved, out string fileText, out string readError))
            {
                failure = DocumentCheckpointMutatingTools.Fail(readError);
                return true;
            }

            ActionContentHelper.Payload payload = ActionContentHelper.TryParseFileBody(fileText, args);
            if (payload == null || !payload.Ok)
            {
                failure = DocumentCheckpointMutatingTools.Fail(
                    payload?.Error ?? ActionFormatHelper.MissingFormatError);
                return true;
            }

            return false;
        }

        private static Dictionary<string, object> BuildReplaceProcessArgs(
            ActionContentHelper.Payload payload,
            Dictionary<string, object> original)
        {
            var detail = new Dictionary<string, object>
            {
                ["original"] = original,
                ["new"] = ActionContentHelper.ToInsertOrNewRaw(payload)
            };
            if (payload != null && !payload.IsRuns && payload.FormatSpec != null && payload.FormatSpec.Ok)
            {
                detail["format"] = new Dictionary<string, object>
                {
                    ["char_format"] = payload.FormatSpec.CharFormat
                };
            }

            return new Dictionary<string, object>
            {
                ["action"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "replace",
                        ["detail"] = detail
                    }
                }
            };
        }
    }
}
