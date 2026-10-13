using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    /// <summary>
    /// 将显式 char_format 应用到 target_codes（charFormatOnly，不改文字）。
    /// inherit_from 已取消；不支持 match_role / set_Style。
    /// 文档触点经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyDocumentFormatTool
    {
        private static readonly string[] FieldsApplied =
        {
            "font_name", "font_size", "font_color", "background_color",
            "is_bold", "has_underline", "has_strikethrough"
        };


        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_document_format"] = (args) =>
            {
                try
                {
                    string logPath = EasyWriteLog.BeginSession("apply_document_format");
                    FormatContextHelper.DbgLog($"会话日志: {logPath}");

                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return Task.FromResult(contractFail);
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_apply_document_format] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return Done(false, resolveError?.Error ?? "无法解析文档渠道");
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_document_format] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string validationError = ValidateModeArgs(args, out string mode,
                        out Dictionary<string, object> charFormat, out Dictionary<string, object> paraFormat);
                    if (validationError != null)
                    {
                        return Done(false, validationError);
                    }

                    if (!TryResolveApplyTargets(
                            args,
                            out List<string> targetCodes,
                            out bool paragraphMode,
                            out string targetError))
                    {
                        return Done(false, targetError);
                    }

                    ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);
                    var formatAnchors = new List<DocumentActionAmbiguityAnchor>
                    {
                        new DocumentActionAmbiguityAnchor
                        {
                            ActionIndex = 0,
                            Field = paragraphMode ? "target_paragraph_codes" : "target_codes",
                            Codes = string.Join(",", targetCodes),
                            TableId = scopeArgs.TableId
                        }
                    };
                    var precheck = InsertTargetLocatorHelper.TryPrecheckTargets(
                        doc,
                        formatAnchors,
                        docHandle.DisallowBackendApi,
                        paragraphMode
                            ? InsertTargetLocatorHelper.LocatePrecheckStream.ParagraphFormatApply
                            : InsertTargetLocatorHelper.LocatePrecheckStream.SentenceFormatApply,
                        scopeArgs,
                        "F_apply_document_format");
                    if (!precheck.Success)
                    {
                        if (precheck.Ambiguity != null && precheck.Ambiguity.HasBlockingError)
                        {
                            return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                                "F_apply_document_format",
                                precheck.Ambiguity));
                        }

                        return Done(false, precheck.Error);
                    }

                    AmbiguityCheckResult ambiguityResult = precheck.Ambiguity;
                    TableScopeIndex tableScope = precheck.TableScope;
                    Dictionary<string, object> snapshot;
                    string responseMode;
                    List<string> explicitFieldsApplied = null;
                    var warnings = new List<string>();

                    snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(charFormat);
                    if (snapshot == null || snapshot.Count == 0)
                    {
                        return Done(false, "须至少提供 char_format 字段；para_format 不会写入 Word");
                    }

                    explicitFieldsApplied = FormatInheritHelper.ListExplicitFieldsApplied(charFormat);
                    responseMode = "explicit";

                    if (paraFormat != null && paraFormat.Count > 0)
                    {
                        warnings.Add("para_format 已忽略（Phase 3 不 set_Style、不写段落对齐）");
                    }

                    var appliedCodes = new List<string>();
                    var skippedCodes = new List<string>();
                    var applyErrors = new List<string>();

                    foreach (string code in targetCodes)
                    {
                        if (paragraphMode)
                        {
                            if (!FormatContextHelper.IsParagraphCode(code))
                            {
                                skippedCodes.Add(code);
                                applyErrors.Add($"{code}: 非 P_ 编码");
                            }
                            else if (!DocumentState.HasParagraphMapping(code))
                            {
                                skippedCodes.Add(code);
                                applyErrors.Add($"{code}: 映射表中不存在");
                            }
                        }
                        else if (!FormatContextHelper.IsSentenceCode(code))
                        {
                            skippedCodes.Add(code);
                            applyErrors.Add($"{code}: 非 S_ 编码");
                        }
                        else if (string.IsNullOrEmpty(DocumentState.GetSentenceContent(code)))
                        {
                            skippedCodes.Add(code);
                            applyErrors.Add($"{code}: 映射表中不存在");
                        }
                    }

                    Word.Range targetSpan = null;
                    if (skippedCodes.Count < targetCodes.Count)
                    {
                        var spanLocate = InsertTargetLocatorHelper.TryResolveCodeSpan(
                            doc,
                            targetCodes,
                            scopeArgs.TableId,
                            InsertTargetLocatorHelper.LocateRangeKind.ContentSpan,
                            paragraphMode,
                            docHandle.DisallowBackendApi);
                        if (spanLocate.Success)
                        {
                            targetSpan = spanLocate.Range;
                        }

                        if (targetSpan == null)
                        {
                            foreach (string code in targetCodes)
                            {
                                if (!skippedCodes.Contains(code))
                                {
                                    skippedCodes.Add(code);
                                    applyErrors.Add($"{code}: 无法在文档中定位序列 Span");
                                }
                            }
                        }
                        else
                        {
                            try
                            {
                                FormatInheritHelper.ApplySnapshot(targetSpan, snapshot, charFormatOnly: true);
                                foreach (string code in targetCodes)
                                {
                                    if (skippedCodes.Contains(code))
                                    {
                                        continue;
                                    }

                                    appliedCodes.Add(code);
                                    FormatContextHelper.DbgLog(
                                        $"apply explicit {code} fields={string.Join(",", explicitFieldsApplied)}");
                                }
                            }
                            catch (Exception ex)
                            {
                                foreach (string code in targetCodes)
                                {
                                    if (!skippedCodes.Contains(code))
                                    {
                                        skippedCodes.Add(code);
                                        applyErrors.Add($"{code}: {ex.Message}");
                                    }
                                }
                            }
                        }
                    }

                    bool success = appliedCodes.Count > 0;
                    var data = new Dictionary<string, object>
                    {
                        ["success"] = success,
                        ["mode"] = responseMode,
                        ["applied_codes"] = appliedCodes,
                        ["fields_applied"] = FieldsApplied.ToList(),
                        ["skipped_codes"] = skippedCodes
                    };

                    if (explicitFieldsApplied != null && explicitFieldsApplied.Count > 0)
                    {
                        data["explicit_fields_applied"] = explicitFieldsApplied;
                    }

                    if (warnings.Count > 0)
                    {
                        data["warning"] = string.Join("; ", warnings);
                    }

                    if (applyErrors.Count > 0)
                    {
                        data["format_apply_error"] = string.Join("; ", applyErrors);
                    }

                    if (!success)
                    {
                        if (appliedCodes.Count == 0)
                        {
                            data["checkpoint_cleanup_only"] = true;
                        }

                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = applyErrors.Count > 0 ? applyErrors[0] : "未能对任何 target 应用格式",
                            Data = data
                        });
                    }

                    Word.Application app = wordApplication as Word.Application;
                    if (app != null && targetSpan != null)
                    {
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, targetSpan, "apply_document_format");
                    }

                    return Task.FromResult(new ToolResult { Success = true, Data = data });
                }
                catch (Exception ex)
                {
                    FormatContextHelper.DbgLog($"F_apply_document_format 异常: {ex.Message}");
                    return Done(false, $"apply_document_format 失败: {ex.Message}");
                }
            };
        }

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            string validationError = ValidateModeArgs(args, out _, out _, out _);
            if (validationError != null)
            {
                failure = DocumentCheckpointMutatingTools.Fail(validationError);
                return true;
            }

            if (!TryResolveApplyTargets(args, out _, out _, out string targetError))
            {
                failure = DocumentCheckpointMutatingTools.Fail(targetError);
                return true;
            }

            return false;
        }

        private static string ValidateModeArgs(
            Dictionary<string, object> args,
            out string mode,
            out Dictionary<string, object> charFormat,
            out Dictionary<string, object> paraFormat)
        {
            mode = "";
            charFormat = null;
            paraFormat = null;

            bool hasSentence = args.ContainsKey("target_codes")
                && !string.IsNullOrWhiteSpace(args["target_codes"]?.ToString());
            bool hasParagraph = args.ContainsKey("target_paragraph_codes")
                && !string.IsNullOrWhiteSpace(args["target_paragraph_codes"]?.ToString());
            if (!hasSentence && !hasParagraph)
            {
                return "缺少 target_codes 或 target_paragraph_codes";
            }

            if (hasSentence && hasParagraph)
            {
                return "target_codes 与 target_paragraph_codes 互斥，请二选一";
            }

            if (ActionFormatHelper.HasCancelledInheritKeys(args))
            {
                return ActionFormatHelper.CancelledInheritFromError;
            }

            if (args.ContainsKey("match_role") && args["match_role"] != null)
            {
                return "本期不支持 match_role；请使用 char_format";
            }

            if (args.ContainsKey("semantic_role") && args["semantic_role"] != null)
            {
                return "本期不支持 match_role / semantic_role；请使用 char_format";
            }

            if (args.ContainsKey("search") && args["search"] != null)
            {
                return "本期不支持 match_role / search；请使用 char_format";
            }

            mode = args.ContainsKey("mode") ? args["mode"]?.ToString()?.Trim() ?? "" : "";

            if (mode == "match_role")
            {
                return "本期不支持 match_role；请使用 char_format";
            }

            if (!string.IsNullOrEmpty(mode) && !string.Equals(mode, "explicit", StringComparison.OrdinalIgnoreCase))
            {
                return $"不支持的 mode: {mode}；仅支持 mode=explicit 或省略并传 char_format";
            }

            mode = "explicit";
            charFormat = ParseNestedObject(args, "char_format");
            paraFormat = ParseNestedObject(args, "para_format");
            if (charFormat == null || charFormat.Count == 0)
            {
                return "必须指定 char_format";
            }

            return null;
        }

        private static bool TryResolveApplyTargets(
            Dictionary<string, object> args,
            out List<string> targetCodes,
            out bool paragraphMode,
            out string error)
        {
            targetCodes = new List<string>();
            paragraphMode = false;
            error = null;

            string paragraphCsv = args.ContainsKey("target_paragraph_codes")
                ? args["target_paragraph_codes"]?.ToString()?.Trim() ?? ""
                : "";
            if (!string.IsNullOrEmpty(paragraphCsv))
            {
                targetCodes = FormatContextHelper.ParseParagraphCodes(paragraphCsv);
                if (targetCodes.Count == 0)
                {
                    error = "target_paragraph_codes 为空";
                    return false;
                }

                paragraphMode = true;
                return true;
            }

            string sentenceCsv = args.ContainsKey("target_codes")
                ? args["target_codes"]?.ToString()?.Trim() ?? ""
                : "";
            targetCodes = FormatContextHelper.ParseSentenceCodes(sentenceCsv);
            if (targetCodes.Count == 0)
            {
                error = "target_codes 为空";
                return false;
            }

            if (targetCodes.Any(c => c.StartsWith("T_", StringComparison.Ordinal)))
            {
                error = "apply 仅支持 S_ / P_，不支持 T_";
                return false;
            }

            bool anyP = targetCodes.Any(FormatContextHelper.IsParagraphCode);
            bool anyS = targetCodes.Any(FormatContextHelper.IsSentenceCode);
            if (anyP && anyS)
            {
                error = "target_codes 不可混用 S_ 与 P_；撞码段改 char 请只传 P_xxxxx_N";
                return false;
            }

            if (anyP)
            {
                paragraphMode = true;
            }

            return true;
        }

        private static Dictionary<string, object> ParseNestedObject(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out object raw) || raw == null)
            {
                return null;
            }

            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            try
            {
                string json = raw.ToString();
                if (string.IsNullOrWhiteSpace(json))
                {
                    return null;
                }

                var serializer = new JavaScriptSerializer();
                var parsed = serializer.DeserializeObject(json);
                if (parsed is Dictionary<string, object> parsedDict)
                {
                    return parsedDict;
                }
            }
            catch
            {
                // fall through
            }

            return null;
        }

        private static Task<ToolResult> Done(bool success, string error)
        {
            return Task.FromResult(new ToolResult { Success = success, Error = error });
        }
    }
}
