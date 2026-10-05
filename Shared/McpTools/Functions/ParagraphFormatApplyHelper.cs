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
            Dictionary<string, object> args,
            Word.Document document = null)
        {
            Word.Document doc = document;
            if (doc == null)
            {
                if (wordApp == null)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = "Word 应用程序不可用" });
                }

                try
                {
                    doc = wordApp.ActiveDocument;
                }
                catch
                {
                    doc = null;
                }
            }

            if (doc == null)
            {
                return Task.FromResult(new ToolResult { Success = false, Error = "没有活动的 Word 文档" });
            }

            if (wordApp == null)
            {
                try
                {
                    wordApp = doc.Application;
                }
                catch
                {
                    wordApp = null;
                }
            }

            DocumentState.BindAndActivate(doc);

            string validationError = ValidateArgs(args, out string mode, out List<string> targetParagraphCodes,
                out Dictionary<string, object> paraFormat, out List<string> warnings);
            if (validationError != null)
            {
                return Task.FromResult(new ToolResult { Success = false, Error = validationError });
            }

            ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);
            AmbiguityCheckResult ambiguityResult = ParagraphFormatAmbiguityHelper.RunPrecheck(
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
            Word.Range targetSpan = null;

            targetSpan = ApplyToTargetParagraphSpan(
                doc,
                targetParagraphCodes,
                ambiguityResult,
                scopeArgs,
                tableScope,
                paraFormat,
                appliedParagraphCodes,
                skippedParagraphCodes,
                applyErrors,
                applyWarnings,
                appliedFieldsUnion);

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

            if (targetSpan != null)
            {
                PostModifyNavigateHelper.NavigateAfterEnd(wordApp, targetSpan, "apply_paragraph_format");
            }

            return Task.FromResult(new ToolResult { Success = true, Data = data });
        }

        private static Word.Range ApplyToTargetParagraphSpan(
            Word.Document doc,
            List<string> targetParagraphCodes,
            AmbiguityCheckResult ambiguityResult,
            ApplyFormatScopeArgs scopeArgs,
            TableScopeIndex tableScope,
            Dictionary<string, object> paraFormat,
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
                }
                else if (string.IsNullOrEmpty(DocumentState.GetParagraphContent(code)))
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 映射表中不存在");
                }
            }

            if (skippedParagraphCodes.Count >= targetParagraphCodes.Count)
            {
                return null;
            }

            int targetDisplayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                ambiguityResult, 0, "target_paragraph_codes");
            if (targetDisplayStart < 0)
            {
                foreach (string code in targetParagraphCodes)
                {
                    if (!skippedParagraphCodes.Contains(code))
                    {
                        skippedParagraphCodes.Add(code);
                        applyErrors.Add($"{code}: 无法在文档中定位序列 Span");
                    }
                }

                return null;
            }

            Word.Range targetSpan = DisplayPositionRangeResolver.ResolveParagraphSpan(
                doc,
                targetParagraphCodes,
                targetDisplayStart,
                scopeArgs.TableId,
                tableScope,
                debugTag: "apply_para_explicit_target_span");
            if (targetSpan == null)
            {
                foreach (string code in targetParagraphCodes)
                {
                    if (!skippedParagraphCodes.Contains(code))
                    {
                        skippedParagraphCodes.Add(code);
                        applyErrors.Add($"{code}: 无法在文档中定位序列 Span");
                    }
                }

                return null;
            }

            try
            {
                List<string> fieldWarnings;
                List<string> appliedFields = ParaFormatWriter.Apply(targetSpan, paraFormat, out fieldWarnings);
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

                foreach (string code in targetParagraphCodes)
                {
                    if (!skippedParagraphCodes.Contains(code))
                    {
                        appliedParagraphCodes.Add(code);
                    }
                }
            }
            catch (Exception ex)
            {
                foreach (string code in targetParagraphCodes)
                {
                    if (!skippedParagraphCodes.Contains(code))
                    {
                        skippedParagraphCodes.Add(code);
                        applyErrors.Add($"{code}: {ex.Message}");
                    }
                }

                return null;
            }

            return targetSpan;
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
            out Dictionary<string, object> paraFormat,
            out List<string> warnings)
        {
            mode = "explicit";
            targetParagraphCodes = new List<string>();
            paraFormat = null;
            warnings = new List<string>();

            ResolveTargetParagraphCodes(args, out targetParagraphCodes, warnings);
            if (targetParagraphCodes.Count == 0)
            {
                return "缺少 target_paragraph_codes 或 target_codes";
            }

            if (ActionFormatHelper.HasCancelledInheritKeys(args))
            {
                return ActionFormatHelper.CancelledInheritFromParagraphError;
            }

            string modeArg = args.ContainsKey("mode") ? args["mode"]?.ToString()?.Trim() ?? "" : "";
            if (!string.IsNullOrEmpty(modeArg)
                && !string.Equals(modeArg, "explicit", StringComparison.OrdinalIgnoreCase))
            {
                return $"不支持的 mode: {modeArg}；仅支持 mode=explicit 或省略并传 para_format";
            }

            paraFormat = ParseNestedObject(args, "para_format");
            if (paraFormat == null || paraFormat.Count == 0)
            {
                return "必须指定 para_format";
            }

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
