using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Phase 2/3：将格式应用到 target_codes（charFormatOnly，不改文字）。
    /// 支持 inherit_from 与 mode=explicit；不支持 match_role / set_Style。
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

                    if (wordApplication == null)
                    {
                        return Done(false, "Word应用程序实例不可用");
                    }

                    dynamic wordApp = wordApplication;
                    if (wordApp.ActiveDocument == null)
                    {
                        return Done(false, "没有活动的Word文档");
                    }

                    DocumentState.BindAndActivate(wordApp.ActiveDocument);

                    string validationError = ValidateModeArgs(args, out string mode, out string inheritFrom,
                        out Dictionary<string, object> charFormat, out Dictionary<string, object> paraFormat);
                    if (validationError != null)
                    {
                        return Done(false, validationError);
                    }

                    string targetCodesCsv = args["target_codes"]?.ToString()?.Trim() ?? "";
                    var targetCodes = FormatContextHelper.ParseSentenceCodes(targetCodesCsv);
                    if (targetCodes.Count == 0)
                    {
                        return Done(false, "target_codes 为空");
                    }

                    if (targetCodes.Any(c => c.StartsWith("T_", StringComparison.Ordinal)))
                    {
                        return Done(false, "apply 仅支持 S_ 句子编码，不支持 T_");
                    }

                    ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);

                    AmbiguityCheckResult ambiguityResult = ApplyFormatAmbiguityHelper.RunPrecheck(
                        mode,
                        inheritFrom,
                        targetCodes,
                        scopeArgs,
                        "F_apply_document_format",
                        out TableScopeIndex tableScope,
                        out string cacheError);
                    if (cacheError != null)
                    {
                        return Done(false, cacheError);
                    }

                    if (ambiguityResult != null && ambiguityResult.HasBlockingError)
                    {
                        return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                            "F_apply_document_format",
                            ambiguityResult));
                    }

                    Word.Document doc = wordApp.ActiveDocument;
                    Dictionary<string, object> snapshot;
                    string responseMode;
                    List<string> explicitFieldsApplied = null;
                    var warnings = new List<string>();

                    if (mode == "explicit")
                    {
                        snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(charFormat);
                        if (snapshot.Count == 0)
                        {
                            return Done(false,
                                "explicit 须至少提供 char_format 字段；para_format 不会写入 Word，请传 char_format 或改用 inherit_from");
                        }

                        explicitFieldsApplied = FormatInheritHelper.ListExplicitFieldsApplied(charFormat);
                        responseMode = "explicit";

                        if (paraFormat != null && paraFormat.Count > 0)
                        {
                            warnings.Add("para_format 已忽略（Phase 3 不 set_Style、不写段落对齐）");
                        }
                    }
                    else
                    {
                        if (!FormatContextHelper.IsSentenceCode(inheritFrom))
                        {
                            return Done(false, "inherit_from 须为 S_ 句子编码；不支持 magic 字符串（如 replaced_range）");
                        }

                        string sourceContent = DocumentState.GetSentenceContent(inheritFrom);
                        if (string.IsNullOrEmpty(sourceContent))
                        {
                            return Done(false, $"inherit_from 不存在: {inheritFrom}");
                        }

                        int inheritDisplayStart = ApplyFormatAmbiguityHelper.GetDisplayStartAt(
                            ambiguityResult, 0, "inherit_from");
                        Word.Range sourceRange = inheritDisplayStart >= 0
                            ? SentenceCodeLocator.LocateRangeByDisplayStart(
                                doc,
                                inheritFrom,
                                inheritDisplayStart,
                                scopeArgs.InheritFromTableId,
                                tableScope,
                                debugTag: $"apply_source:{inheritFrom}")
                            : null;
                        if (sourceRange == null)
                        {
                            return Done(false, $"无法定位 inherit_from: {inheritFrom}");
                        }

                        snapshot = FormatInheritHelper.ExtractSnapshot(sourceRange);
                        responseMode = "inherit";
                    }

                    var appliedCodes = new List<string>();
                    var skippedCodes = new List<string>();
                    var applyErrors = new List<string>();
                    string lastAppliedCode = null;
                    int lastAppliedTargetIndex = -1;

                    foreach (string code in targetCodes)
                    {
                        if (!FormatContextHelper.IsSentenceCode(code))
                        {
                            skippedCodes.Add(code);
                            applyErrors.Add($"{code}: 非 S_ 编码");
                            continue;
                        }

                        string content = DocumentState.GetSentenceContent(code);
                        if (string.IsNullOrEmpty(content))
                        {
                            skippedCodes.Add(code);
                            applyErrors.Add($"{code}: 映射表中不存在");
                            continue;
                        }

                        int targetIndex = targetCodes.IndexOf(code);
                        int targetDisplayStart = ApplyFormatAmbiguityHelper.GetDisplayStartAt(
                            ambiguityResult, targetIndex, "target_codes");
                        Word.Range targetRange = targetDisplayStart >= 0
                            ? SentenceCodeLocator.LocateRangeByDisplayStart(
                                doc,
                                code,
                                targetDisplayStart,
                                scopeArgs.TableId,
                                tableScope,
                                debugTag: $"apply_target:{code}")
                            : null;
                        if (targetRange == null)
                        {
                            skippedCodes.Add(code);
                            applyErrors.Add($"{code}: 无法在文档中定位");
                            continue;
                        }

                        try
                        {
                            FormatInheritHelper.ApplySnapshot(targetRange, snapshot, charFormatOnly: true);
                            appliedCodes.Add(code);
                            lastAppliedCode = code;
                            lastAppliedTargetIndex = targetIndex;
                            if (responseMode == "inherit")
                            {
                                FormatContextHelper.DbgLog(
                                    $"apply {code} <- {inheritFrom} fingerprint={FormatInheritHelper.BuildFingerprint(snapshot)}");
                            }
                            else
                            {
                                FormatContextHelper.DbgLog(
                                    $"apply explicit {code} fields={string.Join(",", explicitFieldsApplied)}");
                            }
                        }
                        catch (Exception ex)
                        {
                            skippedCodes.Add(code);
                            applyErrors.Add($"{code}: {ex.Message}");
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

                    if (responseMode == "inherit")
                    {
                        data["inherit_from"] = inheritFrom;
                    }
                    else if (explicitFieldsApplied != null && explicitFieldsApplied.Count > 0)
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

                    Word.Application app = wordApp as Word.Application;
                    if (app != null && !string.IsNullOrEmpty(lastAppliedCode) && lastAppliedTargetIndex >= 0)
                    {
                        int lastDisplayStart = ApplyFormatAmbiguityHelper.GetDisplayStartAt(
                            ambiguityResult, lastAppliedTargetIndex, "target_codes");
                        Word.Range navRange = lastDisplayStart >= 0
                            ? SentenceCodeLocator.LocateRangeByDisplayStart(
                                doc,
                                lastAppliedCode,
                                lastDisplayStart,
                                scopeArgs.TableId,
                                tableScope,
                                debugTag: $"apply_document_format_nav:{lastAppliedCode}")
                            : null;
                        if (navRange != null)
                        {
                            PostModifyNavigateHelper.NavigateAfterEnd(
                                app, navRange, "apply_document_format");
                        }
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

        private static string ValidateModeArgs(
            Dictionary<string, object> args,
            out string mode,
            out string inheritFrom,
            out Dictionary<string, object> charFormat,
            out Dictionary<string, object> paraFormat)
        {
            mode = "";
            inheritFrom = "";
            charFormat = null;
            paraFormat = null;

            if (!args.ContainsKey("target_codes") || string.IsNullOrWhiteSpace(args["target_codes"]?.ToString()))
            {
                return "缺少 target_codes";
            }

            if (args.ContainsKey("match_role") && args["match_role"] != null)
            {
                return "本期不支持 match_role；请使用 inherit_from 或 mode=explicit + char_format";
            }

            if (args.ContainsKey("semantic_role") && args["semantic_role"] != null)
            {
                return "本期不支持 match_role / semantic_role；请使用 inherit_from 或 mode=explicit + char_format";
            }

            if (args.ContainsKey("search") && args["search"] != null)
            {
                return "本期不支持 match_role / search；请使用 inherit_from 或 mode=explicit + char_format";
            }

            mode = args.ContainsKey("mode") ? args["mode"]?.ToString()?.Trim() ?? "" : "";
            inheritFrom = args.ContainsKey("inherit_from") ? args["inherit_from"]?.ToString()?.Trim() ?? "" : "";

            if (mode == "match_role")
            {
                return "本期不支持 match_role；请使用 inherit_from 或 mode=explicit + char_format";
            }

            if (!string.IsNullOrEmpty(mode) && !string.Equals(mode, "explicit", StringComparison.OrdinalIgnoreCase))
            {
                return $"不支持的 mode: {mode}；仅支持 inherit_from 或 mode=explicit";
            }

            if (string.Equals(mode, "explicit", StringComparison.OrdinalIgnoreCase))
            {
                mode = "explicit";
                if (!string.IsNullOrEmpty(inheritFrom))
                {
                    return "inherit_from 与 mode=explicit 互斥，请只传一种";
                }

                charFormat = ParseNestedObject(args, "char_format");
                paraFormat = ParseNestedObject(args, "para_format");

                bool hasChar = charFormat != null && charFormat.Count > 0;
                bool hasPara = paraFormat != null && paraFormat.Count > 0;
                if (!hasChar && !hasPara)
                {
                    return "mode=explicit 时 char_format 与 para_format 不能同时为空";
                }

                return null;
            }

            if (string.IsNullOrEmpty(inheritFrom))
            {
                return "缺少 inherit_from；或传 mode=explicit + char_format";
            }

            if (args.ContainsKey("char_format") && args["char_format"] != null)
            {
                return "inherit 模式不支持 char_format；请去掉 mode 并使用 inherit_from，或改用 mode=explicit";
            }

            if (args.ContainsKey("para_format") && args["para_format"] != null)
            {
                return "inherit 模式不支持 para_format；请使用 inherit_from 或 mode=explicit";
            }

            mode = "inherit";
            return null;
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
