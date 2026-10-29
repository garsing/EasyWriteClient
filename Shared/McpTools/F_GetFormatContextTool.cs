using System;
using System.Collections;
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

                    // 只读：不 Activate 文档窗口，保持 Desktop 在前台
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError,
                            activateDocument: false))
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

                    return RunFormatContextBatch(args, doc, docHandle);
                }
                catch (Exception ex)
                {
                    FormatContextHelper.DbgLog($"F_get_format_context 异常: {ex.Message}");
                    return Task.FromResult(new ToolResult { Success = false, Error = $"get_format_context 失败: {ex.Message}" });
                }
            };
        }

        private static Task<ToolResult> RunFormatContextBatch(
            Dictionary<string, object> args,
            Word.Document doc,
            InteropDocumentHandle docHandle)
        {
            FormatApplyBatchCompiler.CompileResult compiled = FormatApplyBatchCompiler.CompileContext(args);
            if (!string.IsNullOrEmpty(compiled.ContractError))
            {
                return Task.FromResult(new ToolResult { Success = false, Error = compiled.ContractError });
            }

            bool passthrough = !FormatApplyBatchCompiler.HasItems(args)
                && compiled.Ops.Count == 1
                && compiled.Warnings.Count == 0;
            if (passthrough)
            {
                return ReadOne(args, doc, docHandle);
            }

            var entries = new List<object>();
            var warnings = new List<object>();
            foreach (Dictionary<string, object> warning in compiled.Warnings)
            {
                warnings.Add(warning);
            }

            foreach (FormatApplyBatchCompiler.CompiledOp op in compiled.Ops)
            {
                Dictionary<string, object> oneArgs = FormatApplyBatchCompiler.CopyArgsForOp(args, op, null);
                ToolResult one = ReadOne(oneArgs, doc, docHandle).GetAwaiter().GetResult();
                string codes = string.Join(",", op.Codes ?? new List<string>());
                if (one != null && one.Success)
                {
                    foreach (Dictionary<string, object> entry in ReadEntries(one.Data))
                    {
                        entry["item_index"] = op.ItemIndex;
                        entries.Add(entry);
                    }
                }
                else
                {
                    warnings.Add(FormatApplyBatchCompiler.Warning(
                        op.ItemIndex,
                        codes,
                        one?.Error ?? "无法定位"));
                }
            }

            var data = new Dictionary<string, object>
            {
                ["context_window"] = 0,
                ["entries"] = entries,
                ["warnings"] = warnings
            };
            if (entries.Count == 0)
            {
                string error = "未能读取任何格式";
                if (warnings.Count > 0 && warnings[0] is Dictionary<string, object> first)
                {
                    error = first.ContainsKey("message") ? first["message"]?.ToString() ?? error : error;
                }

                return Task.FromResult(new ToolResult
                {
                    Success = false,
                    Error = error,
                    Data = data
                });
            }

            FormatContextHelper.DbgLog(
                $"F_get_format_context 批量成功 entries={entries.Count} warnings={warnings.Count}");
            return Task.FromResult(new ToolResult { Success = true, Data = data });
        }

        private static IEnumerable<Dictionary<string, object>> ReadEntries(object data)
        {
            var dict = data as Dictionary<string, object>;
            if (dict == null || !dict.TryGetValue("entries", out object raw) || raw == null)
            {
                yield break;
            }

            if (raw is IEnumerable enumerable && !(raw is string))
            {
                foreach (object item in enumerable)
                {
                    Dictionary<string, object> entry = item as Dictionary<string, object>
                        ?? ActionFormatHelper.CoerceDict(item);
                    if (entry != null)
                    {
                        yield return entry;
                    }
                }
            }
        }

        private static Task<ToolResult> ReadOne(
            Dictionary<string, object> args,
            Word.Document doc,
            InteropDocumentHandle docHandle)
        {
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
                        var paraAnchors = new List<DocumentActionAmbiguityAnchor>
                        {
                            new DocumentActionAmbiguityAnchor
                            {
                                ActionIndex = 0,
                                Field = "target_paragraph_codes",
                                Codes = string.Join(",", parsedParagraphTargets),
                                TableId = scopeArgs.TableId
                            }
                        };
                        var paraPrecheck = InsertTargetLocatorHelper.TryPrecheckTargets(
                            doc,
                            paraAnchors,
                            docHandle.DisallowBackendApi,
                            InsertTargetLocatorHelper.LocatePrecheckStream.ParagraphFormatContext,
                            scopeArgs,
                            "F_get_format_context");
                        if (!paraPrecheck.Success)
                        {
                            if (paraPrecheck.Ambiguity != null && paraPrecheck.Ambiguity.HasBlockingError)
                            {
                                return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                                    "F_get_format_context",
                                    paraPrecheck.Ambiguity));
                            }

                            return Task.FromResult(new ToolResult { Success = false, Error = paraPrecheck.Error });
                        }

                        built = FormatContextHelper.BuildParagraphContext(
                            doc,
                            targetParagraphCodes,
                            contextWindow,
                            scopeArgs.TableId,
                            paraPrecheck.TableScope,
                            paraPrecheck.Ambiguity,
                            docHandle.DisallowBackendApi);
                    }
                    else
                    {
                        string targetCodes = args["target_codes"].ToString()?.Trim() ?? "";
                        string formatPolicy = FormatContextHelper.LocateThenFormatPolicyError(
                            doc,
                            targetCodes,
                            docHandle.DisallowBackendApi);
                        if (!string.IsNullOrEmpty(formatPolicy))
                        {
                            return Task.FromResult(new ToolResult
                            {
                                Success = false,
                                Error = formatPolicy
                            });
                        }

                        var parsedTargets = FormatContextHelper.ParseSentenceCodes(targetCodes);
                        ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);
                        var sentAnchors = new List<DocumentActionAmbiguityAnchor>
                        {
                            new DocumentActionAmbiguityAnchor
                            {
                                ActionIndex = 0,
                                Field = "target_codes",
                                Codes = string.Join(",", parsedTargets),
                                TableId = scopeArgs.TableId
                            }
                        };
                        var sentPrecheck = InsertTargetLocatorHelper.TryPrecheckTargets(
                            doc,
                            sentAnchors,
                            docHandle.DisallowBackendApi,
                            InsertTargetLocatorHelper.LocatePrecheckStream.SentenceFormatContext,
                            scopeArgs,
                            "F_get_format_context");
                        if (!sentPrecheck.Success)
                        {
                            if (sentPrecheck.Ambiguity != null && sentPrecheck.Ambiguity.HasBlockingError)
                            {
                                return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                                    "F_get_format_context",
                                    sentPrecheck.Ambiguity));
                            }

                            return Task.FromResult(new ToolResult { Success = false, Error = sentPrecheck.Error });
                        }

                        built = FormatContextHelper.BuildContext(
                            doc,
                            targetCodes,
                            contextWindow,
                            scopeArgs.TableId,
                            sentPrecheck.TableScope,
                            sentPrecheck.Ambiguity,
                            docHandle.DisallowBackendApi);
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
    }
}
