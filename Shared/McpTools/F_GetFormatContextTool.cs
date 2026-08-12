using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Phase 2：读取 target 及邻居的格式上下文（只读）。经 DocumentHost 解析渠道（Word/WPS）。
    /// </summary>
    public static class F_GetFormatContextTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_format_context"] = (args) =>
            {
                try
                {
                    string logPath = EasyWriteLog.BeginSession("get_format_context");
                    FormatContextHelper.DbgLog($"会话日志: {logPath}");

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_get_format_context] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return Task.FromResult(resolveError);
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_get_format_context] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    bool hasTargetCodes = args.ContainsKey("target_codes") && args["target_codes"] != null
                        && !string.IsNullOrWhiteSpace(args["target_codes"].ToString());
                    bool hasTargetParagraphCodes = args.ContainsKey("target_paragraph_codes")
                        && args["target_paragraph_codes"] != null
                        && !string.IsNullOrWhiteSpace(args["target_paragraph_codes"].ToString());

                    if (hasTargetCodes && hasTargetParagraphCodes)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "target_codes 与 target_paragraph_codes 互斥，请二选一（mutually_exclusive_targets）"
                        });
                    }

                    if (!hasTargetCodes && !hasTargetParagraphCodes)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "缺少 target_codes 或 target_paragraph_codes（missing_target）"
                        });
                    }

                    int contextWindow = FormatContextHelper.DefaultContextWindow;
                    if (args.ContainsKey("context_window") && args["context_window"] != null)
                    {
                        if (!int.TryParse(args["context_window"].ToString(), out contextWindow) || contextWindow < 0)
                        {
                            return Task.FromResult(new ToolResult { Success = false, Error = "context_window 须为非负整数" });
                        }
                    }

                    FormatContextHelper.FormatContextBuildResult built;
                    if (hasTargetParagraphCodes)
                    {
                        string targetParagraphCodes = args["target_paragraph_codes"].ToString()?.Trim() ?? "";
                        var parsedParagraphTargets = FormatContextHelper.ParseParagraphCodes(targetParagraphCodes);
                        ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);
                        AmbiguityCheckResult precheck = ParagraphFormatAmbiguityHelper.RunContextReadPrecheck(
                            parsedParagraphTargets,
                            scopeArgs,
                            "F_get_format_context",
                            out TableScopeIndex tableScope,
                            out string cacheError);
                        if (cacheError != null)
                        {
                            return Task.FromResult(new ToolResult { Success = false, Error = cacheError });
                        }

                        if (precheck != null && precheck.HasBlockingError)
                        {
                            return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                                "F_get_format_context",
                                precheck));
                        }

                        built = FormatContextHelper.BuildParagraphContext(
                            doc,
                            targetParagraphCodes,
                            contextWindow,
                            scopeArgs.TableId,
                            tableScope,
                            precheck);
                    }
                    else
                    {
                        string targetCodes = args["target_codes"].ToString()?.Trim() ?? "";
                        var parsedTargets = FormatContextHelper.ParseSentenceCodes(targetCodes);
                        ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);
                        AmbiguityCheckResult precheck = ApplyFormatAmbiguityHelper.RunContextReadPrecheck(
                            parsedTargets,
                            scopeArgs,
                            "F_get_format_context",
                            out TableScopeIndex tableScope,
                            out string cacheError);
                        if (cacheError != null)
                        {
                            return Task.FromResult(new ToolResult { Success = false, Error = cacheError });
                        }

                        if (precheck != null && precheck.HasBlockingError)
                        {
                            return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                                "F_get_format_context",
                                precheck));
                        }

                        built = FormatContextHelper.BuildContext(
                            doc,
                            targetCodes,
                            contextWindow,
                            scopeArgs.TableId,
                            tableScope,
                            precheck);
                    }

                    if (!built.Success)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = built.Error ?? "get_format_context 失败",
                            Data = built.Errors.Count > 0
                                ? new Dictionary<string, object> { ["errors"] = built.Errors }
                                : null
                        });
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["context_window"] = built.ContextWindow,
                        ["entries"] = built.Entries
                    };

                    if (built.IsParagraphMode)
                    {
                        data["target_paragraph_codes"] = built.TargetParagraphCodes;
                    }
                    else
                    {
                        data["target_codes"] = built.TargetCodes;
                    }

                    if (built.Errors.Count > 0)
                    {
                        data["errors"] = built.Errors;
                    }

                    FormatContextHelper.DbgLog(
                        $"F_get_format_context 成功 entries={built.Entries.Count} " +
                        $"paragraphMode={built.IsParagraphMode}");
                    return Task.FromResult(new ToolResult { Success = true, Data = data });
                }
                catch (Exception ex)
                {
                    FormatContextHelper.DbgLog($"F_get_format_context 异常: {ex.Message}");
                    return Task.FromResult(new ToolResult { Success = false, Error = $"get_format_context 失败: {ex.Message}" });
                }
            };
        }
    }
}
