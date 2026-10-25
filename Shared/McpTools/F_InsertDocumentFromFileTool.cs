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
                    ActionContentHelper.Payload payload = ActionContentHelper.TryParseFileBody(fileText, args);

                    var build = BuildInsertProcessArgs(payload, args);
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

                    EasyWriteDiagnostics.Log(
                        DebugCategory.Tool,
                        "[F_insert_document_from_file] delegate → F_process_document_actions; channel_id="
                        + (ChannelContext.TryGetChannelIdFromParameters(build.args) ?? "(default)"));

                    ToolResult processResult = await processHandler(build.args);
                    if (!processResult.Success)
                    {
                        WordAddIn1.DocumentHost.HostProbe.Log(
                            "from-file-ins-fail",
                            "err=" + (processResult.Error ?? "(null)")
                            + " data=" + BriefData(processResult.Data));
                        return processResult;
                    }

                    Dictionary<string, object> targetEcho = InsertTargetLocatorHelper.TryGetTargetDict(args);
                    bool start = targetEcho != null
                        && targetEcho.ContainsKey("position")
                        && InsertTargetLocatorHelper.IsStartPosition(targetEcho["position"]?.ToString());
                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            path = resolved.Display,
                            insert_file = resolved.Display,
                            target = targetEcho,
                            characters_inserted = fileText?.Length ?? 0,
                            process_result = processResult.Data,
                            message = start
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

        private static string BriefData(object data)
        {
            if (data == null)
            {
                return "(null)";
            }

            try
            {
                string json = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(data);
                return json.Length > 500 ? json.Substring(0, 500) : json;
            }
            catch (Exception ex)
            {
                return data.GetType().Name + " " + ex.Message;
            }
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

            var build = BuildInsertProcessArgs(payload, args);
            if (!string.IsNullOrEmpty(build.error))
            {
                failure = DocumentCheckpointMutatingTools.Fail(build.error);
                return true;
            }

            return false;
        }

        private static (Dictionary<string, object> args, string error) BuildInsertProcessArgs(
            ActionContentHelper.Payload payload,
            Dictionary<string, object> args)
        {
            string retired = InsertTargetLocatorHelper.RetiredFlatLocatorError(
                args, InsertTargetLocatorHelper.FromFileRetiredInsertKeys);
            if (retired != null)
            {
                return (null, retired);
            }

            var detail = new Dictionary<string, object>
            {
                ["insert"] = ActionContentHelper.ToInsertOrNewRaw(payload)
            };
            if (payload != null && !payload.IsRuns && payload.FormatSpec != null && payload.FormatSpec.Ok)
            {
                detail["format"] = new Dictionary<string, object>
                {
                    ["char_format"] = payload.FormatSpec.CharFormat
                };
            }

            Dictionary<string, object> targetDict = InsertTargetLocatorHelper.TryGetTargetDict(args);
            if (targetDict != null)
            {
                var parsed = InsertTargetLocatorHelper.ParseLocatorObject(
                    targetDict, sentenceOnly: false, fieldLabel: "target");
                if (!parsed.Success)
                {
                    return (null, parsed.Error);
                }

                detail["target"] = targetDict;
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
