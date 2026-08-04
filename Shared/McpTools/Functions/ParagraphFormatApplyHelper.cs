using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    public static class ParagraphFormatApplyHelper
    {
        public static Task<ToolResult> RunApplyAsync(
            Word.Application wordApp,
            Dictionary<string, object> args)
        {
            if (wordApp == null)
            {
                return Task.FromResult(new ToolResult { Success = false, Error = "Word 应用程序不可用" });
            }

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                doc = null;
            }

            if (doc == null)
            {
                return Task.FromResult(new ToolResult { Success = false, Error = "没有活动的 Word 文档" });
            }

            DocumentState.BindAndActivate(doc);

            string validationError = ValidateArgs(args, out string mode, out List<string> targetParagraphCodes,
                out string inheritFromParagraph, out Dictionary<string, object> paraFormat, out List<string> warnings);
            if (validationError != null)
            {
                return Task.FromResult(new ToolResult { Success = false, Error = validationError });
            }

            ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);
            AmbiguityCheckResult ambiguityResult = ParagraphFormatAmbiguityHelper.RunPrecheck(
                inheritFromParagraph,
                scopeArgs,
                targetParagraphCodes,
                "F_apply_paragraph_format",
                out TableScopeIndex tableScope,
                out string cacheError);
            if (cacheError != null)
            {
                return Task.FromResult(new ToolResult { Success = false, Error = cacheError });
            }

            if (ambiguityResult != null && ambiguityResult.HasBlockingError)
            {
                return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                    "F_apply_paragraph_format",
                    ambiguityResult));
            }

            var appliedParagraphCodes = new List<string>();
            var skippedParagraphCodes = new List<string>();
            var applyErrors = new List<string>();
            var applyWarnings = new List<string>(warnings);
            var appliedFieldsUnion = new List<string>();
            string lastAppliedCode = null;
            int lastAppliedIndex = -1;

            if (string.Equals(mode, "inherit", StringComparison.OrdinalIgnoreCase))
            {
                int inheritDisplayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, 0, "inherit_from_paragraph");
                Word.Range sourceRange = inheritDisplayStart >= 0
                    ? ParagraphCodeResolver.ResolveParagraphRangeByDisplayStart(
                        doc,
                        inheritFromParagraph,
                        inheritDisplayStart,
                        scopeArgs.InheritFromTableId,
                        tableScope,
                        debugTag: $"apply_para_inherit_source:{inheritFromParagraph}")
                    : null;
                if (sourceRange == null)
                {
                    return Task.FromResult(new ToolResult
                    {
                        Success = false,
                        Error = $"无法定位 inherit_from_paragraph: {inheritFromParagraph}"
                    });
                }

                ApplyInheritToTargets(
                    doc,
                    sourceRange,
                    targetParagraphCodes,
                    ambiguityResult,
                    scopeArgs,
                    tableScope,
                    appliedParagraphCodes,
                    skippedParagraphCodes,
                    applyErrors,
                    ref lastAppliedCode,
                    ref lastAppliedIndex);
            }
            else
            {
                ApplyExplicitToTargets(
                    doc,
                    paraFormat,
                    targetParagraphCodes,
                    ambiguityResult,
                    scopeArgs,
                    tableScope,
                    appliedParagraphCodes,
                    skippedParagraphCodes,
                    applyErrors,
                    applyWarnings,
                    appliedFieldsUnion);
                if (appliedParagraphCodes.Count > 0)
                {
                    lastAppliedCode = appliedParagraphCodes[appliedParagraphCodes.Count - 1];
                    lastAppliedIndex = targetParagraphCodes.IndexOf(lastAppliedCode);
                }
            }

            var data = BuildResponseData(
                mode,
                appliedParagraphCodes,
                skippedParagraphCodes,
                applyWarnings,
                appliedFieldsUnion,
                applyErrors);

            if (appliedParagraphCodes.Count == 0)
            {
                return Task.FromResult(new ToolResult
                {
                    Success = false,
                    Error = applyErrors.Count > 0 ? applyErrors[0] : "未能对任何 target 应用段落格式",
                    Data = data
                });
            }

            if (!string.IsNullOrEmpty(lastAppliedCode) && lastAppliedIndex >= 0)
            {
                int lastDisplayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, lastAppliedIndex, "target_paragraph_codes");
                Word.Range navRange = lastDisplayStart >= 0
                    ? ParagraphCodeResolver.ResolveParagraphRangeByDisplayStart(
                        doc,
                        lastAppliedCode,
                        lastDisplayStart,
                        scopeArgs.TableId,
                        tableScope,
                        debugTag: $"apply_paragraph_format_nav:{lastAppliedCode}")
                    : null;
                if (navRange != null)
                {
                    PostModifyNavigateHelper.NavigateAfterEnd(wordApp, navRange, "apply_paragraph_format");
                }
            }

            return Task.FromResult(new ToolResult { Success = true, Data = data });
        }

        private static void ApplyInheritToTargets(
            Word.Document doc,
            Word.Range sourceRange,
            List<string> targetParagraphCodes,
            AmbiguityCheckResult ambiguityResult,
            ApplyFormatScopeArgs scopeArgs,
            TableScopeIndex tableScope,
            List<string> appliedParagraphCodes,
            List<string> skippedParagraphCodes,
            List<string> applyErrors,
            ref string lastAppliedCode,
            ref int lastAppliedIndex)
        {
            foreach (string code in targetParagraphCodes)
            {
                if (!FormatContextHelper.IsParagraphCode(code))
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 非 P_ 编码");
                    continue;
                }

                int targetIndex = targetParagraphCodes.IndexOf(code);
                int targetDisplayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, targetIndex, "target_paragraph_codes");
                Word.Range targetRange = targetDisplayStart >= 0
                    ? ParagraphCodeResolver.ResolveParagraphRangeByDisplayStart(
                        doc,
                        code,
                        targetDisplayStart,
                        scopeArgs.TableId,
                        tableScope,
                        debugTag: $"apply_para_inherit_target:{code}")
                    : null;
                if (targetRange == null)
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 无法在文档中定位");
                    continue;
                }

                try
                {
                    ParagraphFormatInheritHelper.CopyParagraphFormat(sourceRange, targetRange);
                    appliedParagraphCodes.Add(code);
                    lastAppliedCode = code;
                    lastAppliedIndex = targetIndex;
                }
                catch (Exception ex)
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: {ex.Message}");
                }
            }
        }

        private static void ApplyExplicitToTargets(
            Word.Document doc,
            Dictionary<string, object> paraFormat,
            List<string> targetParagraphCodes,
            AmbiguityCheckResult ambiguityResult,
            ApplyFormatScopeArgs scopeArgs,
            TableScopeIndex tableScope,
            List<string> appliedParagraphCodes,
            List<string> skippedParagraphCodes,
            List<string> applyErrors,
            List<string> applyWarnings,
            List<string> appliedFieldsUnion)
        {
            foreach (string code in targetParagraphCodes)
            {
                if (!FormatContextHelper.IsParagraphCode(code))
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 非 P_ 编码");
                    continue;
                }

                int targetIndex = targetParagraphCodes.IndexOf(code);
                int targetDisplayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, targetIndex, "target_paragraph_codes");
                Word.Range targetRange = targetDisplayStart >= 0
                    ? ParagraphCodeResolver.ResolveParagraphRangeByDisplayStart(
                        doc,
                        code,
                        targetDisplayStart,
                        scopeArgs.TableId,
                        tableScope,
                        debugTag: $"apply_para_explicit_target:{code}")
                    : null;
                if (targetRange == null)
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 无法在文档中定位");
                    continue;
                }

                try
                {
                    List<string> fieldWarnings;
                    List<string> appliedFields = ParaFormatWriter.Apply(targetRange, paraFormat, out fieldWarnings);
                    foreach (string w in fieldWarnings)
                    {
                        if (!applyWarnings.Contains(w))
                        {
                            applyWarnings.Add(w);
                        }
                    }

                    foreach (string f in appliedFields)
                    {
                        if (!appliedFieldsUnion.Contains(f))
                        {
                            appliedFieldsUnion.Add(f);
                        }
                    }

                    appliedParagraphCodes.Add(code);
                }
                catch (Exception ex)
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: {ex.Message}");
                }
            }
        }

        private static Dictionary<string, object> BuildResponseData(
            string mode,
            List<string> appliedParagraphCodes,
            List<string> skippedParagraphCodes,
            List<string> warnings,
            List<string> appliedFields,
            List<string> applyErrors)
        {
            var sentenceCodes = new List<string>();
            foreach (string pCode in appliedParagraphCodes)
            {
                foreach (string sCode in ParagraphCodeResolver.ExpandParagraphToSentenceCodes(pCode))
                {
                    if (!sentenceCodes.Contains(sCode))
                    {
                        sentenceCodes.Add(sCode);
                    }
                }
            }

            var data = new Dictionary<string, object>
            {
                ["mode"] = mode,
                ["affected_paragraph_codes"] = appliedParagraphCodes,
                ["sentence_codes_in_paragraph"] = sentenceCodes,
                ["skipped_paragraph_codes"] = skippedParagraphCodes,
            };

            if (appliedFields.Count > 0)
            {
                data["applied_fields"] = appliedFields;
            }

            if (warnings.Count > 0)
            {
                data["warnings"] = warnings;
            }

            if (applyErrors.Count > 0)
            {
                data["format_apply_error"] = string.Join("; ", applyErrors);
            }

            return data;
        }

        private static string ValidateArgs(
            Dictionary<string, object> args,
            out string mode,
            out List<string> targetParagraphCodes,
            out string inheritFromParagraph,
            out Dictionary<string, object> paraFormat,
            out List<string> warnings)
        {
            mode = "";
            targetParagraphCodes = new List<string>();
            inheritFromParagraph = "";
            paraFormat = null;
            warnings = new List<string>();

            ResolveTargetParagraphCodes(args, out targetParagraphCodes, warnings);
            if (targetParagraphCodes.Count == 0)
            {
                return "缺少 target_paragraph_codes 或 target_codes";
            }

            mode = args.ContainsKey("mode") ? args["mode"]?.ToString()?.Trim() ?? "" : "";
            inheritFromParagraph = args.ContainsKey("inherit_from_paragraph")
                ? args["inherit_from_paragraph"]?.ToString()?.Trim() ?? ""
                : "";

            if (string.IsNullOrEmpty(inheritFromParagraph)
                && args.ContainsKey("inherit_from")
                && args["inherit_from"] != null)
            {
                string inheritFromSentence = args["inherit_from"].ToString()?.Trim();
                inheritFromParagraph = ParagraphCodeResolver.ResolveParagraphFromSentence(inheritFromSentence) ?? "";
                if (string.IsNullOrEmpty(inheritFromParagraph))
                {
                    return $"inherit_from 无法反查 P_: {inheritFromSentence}";
                }

                warnings.Add("deprecated_inherit_from_sentence_resolved_to_paragraph");
            }

            if (!string.IsNullOrEmpty(mode) && !string.Equals(mode, "explicit", StringComparison.OrdinalIgnoreCase))
            {
                return $"不支持的 mode: {mode}；仅支持 inherit 或 mode=explicit";
            }

            if (string.Equals(mode, "explicit", StringComparison.OrdinalIgnoreCase))
            {
                mode = "explicit";
                if (!string.IsNullOrEmpty(inheritFromParagraph))
                {
                    return "inherit_from_paragraph 与 mode=explicit 互斥";
                }

                paraFormat = ParseNestedObject(args, "para_format");
                if (paraFormat == null || paraFormat.Count == 0)
                {
                    return "mode=explicit 时 para_format 不能为空";
                }

                return null;
            }

            if (string.IsNullOrEmpty(inheritFromParagraph))
            {
                return "缺少 inherit_from_paragraph；或传 mode=explicit + para_format";
            }

            if (args.ContainsKey("para_format") && args["para_format"] != null)
            {
                return "inherit 模式不支持 para_format";
            }

            mode = "inherit";
            return null;
        }

        private static void ResolveTargetParagraphCodes(
            Dictionary<string, object> args,
            out List<string> targetParagraphCodes,
            List<string> warnings)
        {
            targetParagraphCodes = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            string paragraphCsv = args.ContainsKey("target_paragraph_codes")
                ? args["target_paragraph_codes"]?.ToString()?.Trim()
                : "";
            if (!string.IsNullOrEmpty(paragraphCsv))
            {
                foreach (string code in FormatContextHelper.ParseParagraphCodes(paragraphCsv))
                {
                    if (seen.Add(code))
                    {
                        targetParagraphCodes.Add(code);
                    }
                }

                return;
            }

            string sentenceCsv = args.ContainsKey("target_codes")
                ? args["target_codes"]?.ToString()?.Trim()
                : "";
            if (string.IsNullOrEmpty(sentenceCsv))
            {
                return;
            }

            warnings.Add("deprecated_target_codes_resolved_to_paragraph");
            foreach (string sentenceCode in FormatContextHelper.ParseSentenceCodes(sentenceCsv))
            {
                string paragraphCode = ParagraphCodeResolver.ResolveParagraphFromSentence(sentenceCode);
                if (string.IsNullOrEmpty(paragraphCode))
                {
                    continue;
                }

                if (seen.Add(paragraphCode))
                {
                    targetParagraphCodes.Add(paragraphCode);
                }
            }
        }

        private static Dictionary<string, object> ParseNestedObject(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return new Dictionary<string, object>();
            }

            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            try
            {
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                return serializer.DeserializeObject(raw.ToString()) as Dictionary<string, object>
                    ?? new Dictionary<string, object>();
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }
    }
}
