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
            Word.Document document = null,
            bool disallowBackendApi = false)
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
                out List<string> sentenceCodes, out Dictionary<string, object> paraFormat, out List<string> warnings);
            if (validationError != null)
            {
                return Task.FromResult(new ToolResult { Success = false, Error = validationError });
            }

            ApplyFormatScopeArgs scopeArgs = ApplyFormatAmbiguityHelper.ParseScopeArgs(args);
            if (!TextSpanArg.TryRead(args, out TextSpan earlyTextSpan, out string earlyTextError))
            {
                return Task.FromResult(new ToolResult { Success = false, Error = earlyTextError });
            }

            bool sentenceInput = sentenceCodes.Count > 0;
            List<string> locateCodes = sentenceInput ? sentenceCodes : targetParagraphCodes;
            InsertTargetLocatorHelper.LocateSpanBatchResult locatedBatch = InsertTargetLocatorHelper.LocateFormatSpan(
                doc,
                locateCodes,
                paragraph: !sentenceInput,
                scopeArgs.TableId,
                earlyTextSpan,
                disallowBackendApi,
                InsertTargetLocatorHelper.LocatePrecheckStream.ParagraphFormatApply,
                scopeArgs,
                "F_apply_paragraph_format");
            if (locatedBatch.BatchBlocked)
            {
                if (locatedBatch.Ambiguity != null && locatedBatch.Ambiguity.HasBlockingError)
                {
                    return Task.FromResult(SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                        "F_apply_paragraph_format",
                        locatedBatch.Ambiguity));
                }

                return Task.FromResult(new ToolResult { Success = false, Error = locatedBatch.Error });
            }

            AmbiguityCheckResult ambiguityResult = locatedBatch.Ambiguity;
            TableScopeIndex tableScope = locatedBatch.TableScope;
            InsertTargetLocatorHelper.LocateSpanItem locatedItem = locatedBatch.Items[0];

            var appliedParagraphCodes = new List<string>();
            var skippedParagraphCodes = new List<string>();
            var applyErrors = new List<string>();
            var applyWarnings = new List<string>(warnings);
            var formatWarnings = new List<Dictionary<string, object>>();
            var appliedFieldsUnion = new List<string>();
            Word.Range targetSpan = null;

            if (!TextSpanArg.TryRead(args, out TextSpan textSpan, out string textError))
            {
                return Task.FromResult(new ToolResult { Success = false, Error = textError });
            }

            if (!locatedItem.Success && TextSpanArg.IsDirectError(locatedItem.Error))
            {
                return Task.FromResult(new ToolResult { Success = false, Error = locatedItem.Error });
            }

            targetSpan = ApplyToTargetParagraphSpan(
                doc,
                locateCodes,
                sentenceInput,
                scopeArgs,
                paraFormat,
                appliedParagraphCodes,
                skippedParagraphCodes,
                applyErrors,
                applyWarnings,
                formatWarnings,
                appliedFieldsUnion,
                disallowBackendApi,
                textSpan,
                locatedItem.Success ? locatedItem.Range : null);

            var data = BuildResponseData(
                mode,
                appliedParagraphCodes,
                skippedParagraphCodes,
                applyWarnings,
                formatWarnings,
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
            List<string> targetCodes,
            bool sentenceInput,
            ApplyFormatScopeArgs scopeArgs,
            Dictionary<string, object> paraFormat,
            List<string> appliedParagraphCodes,
            List<string> skippedParagraphCodes,
            List<string> applyErrors,
            List<string> applyWarnings,
            List<Dictionary<string, object>> formatWarnings,
            List<string> appliedFieldsUnion,
            bool disallowBackendApi,
            TextSpan textSpan,
            Word.Range presetSpan)
        {
            foreach (string code in targetCodes)
            {
                if (!sentenceInput && !FormatContextHelper.IsParagraphCode(code))
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 非 P_ 编码");
                }
                else if (!sentenceInput && string.IsNullOrEmpty(DocumentState.GetParagraphContent(code)))
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 映射表中不存在");
                }
            }

            if (!sentenceInput && skippedParagraphCodes.Count >= targetCodes.Count)
            {
                return null;
            }

            Word.Range targetSpan = presetSpan;
            if (targetSpan == null)
            {
                foreach (string code in targetCodes)
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
                List<Dictionary<string, object>> fieldFormatWarnings;
                List<string> appliedFields = sentenceInput
                    ? ApplyTouchedParagraphs(targetSpan, paraFormat, out fieldWarnings, out fieldFormatWarnings)
                    : ParaFormatWriter.Apply(
                    targetSpan, paraFormat, out fieldWarnings, out fieldFormatWarnings);
                if (fieldFormatWarnings != null)
                {
                    formatWarnings.AddRange(fieldFormatWarnings);
                }
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

                foreach (string code in targetCodes)
                {
                    if (!skippedParagraphCodes.Contains(code))
                    {
                        appliedParagraphCodes.Add(code);
                    }
                }
            }
            catch (Exception ex)
            {
                foreach (string code in targetCodes)
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

        private static List<string> ApplyTouchedParagraphs(
            Word.Range range,
            Dictionary<string, object> paraFormat,
            out List<string> warnings,
            out List<Dictionary<string, object>> formatWarnings)
        {
            warnings = new List<string>();
            formatWarnings = new List<Dictionary<string, object>>();
            var applied = new List<string>();
            var seen = new HashSet<int>();
            if (range?.Paragraphs == null)
            {
                return applied;
            }

            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                if (paragraph?.Range == null)
                {
                    continue;
                }

                if (paragraph.Range.Start >= range.End || paragraph.Range.End <= range.Start)
                {
                    continue;
                }

                if (!seen.Add(paragraph.Range.Start))
                {
                    continue;
                }

                List<string> oneWarnings;
                List<Dictionary<string, object>> oneFormats;
                List<string> oneApplied = ParaFormatWriter.Apply(
                    paragraph.Range, paraFormat, out oneWarnings, out oneFormats);
                if (oneWarnings != null)
                {
                    foreach (string warning in oneWarnings)
                    {
                        if (!warnings.Contains(warning))
                        {
                            warnings.Add(warning);
                        }
                    }
                }

                if (oneFormats != null)
                {
                    formatWarnings.AddRange(oneFormats);
                }

                if (oneApplied != null)
                {
                    foreach (string field in oneApplied)
                    {
                        if (!applied.Contains(field))
                        {
                            applied.Add(field);
                        }
                    }
                }
            }

            return applied;
        }

        private static Dictionary<string, object> BuildResponseData(
            string mode,
            List<string> appliedParagraphCodes,
            List<string> skippedParagraphCodes,
            List<string> warnings,
            List<Dictionary<string, object>> formatWarnings,
            List<string> appliedFields,
            List<string> applyErrors)
        {
            var data = new Dictionary<string, object>
            {
                ["mode"] = mode,
                ["affected_paragraph_codes"] = appliedParagraphCodes,
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

            if (formatWarnings != null && formatWarnings.Count > 0)
            {
                data["format_warnings"] = formatWarnings;
            }

            if (applyErrors.Count > 0)
            {
                data["format_apply_error"] = string.Join("; ", applyErrors);
            }

            return data;
        }

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            if (FormatApplyBatchCompiler.HasItems(args))
            {
                string itemError = FormatApplyBatchCompiler.ValidateParagraphItems(args);
                if (itemError != null)
                {
                    failure = DocumentCheckpointMutatingTools.Fail(itemError);
                    return true;
                }

                return false;
            }

            string validationError = ValidateArgs(args, out _, out _, out _, out _, out _);
            if (validationError != null)
            {
                failure = DocumentCheckpointMutatingTools.Fail(validationError);
                return true;
            }

            return false;
        }

        private static string ValidateArgs(
            Dictionary<string, object> args,
            out string mode,
            out List<string> targetParagraphCodes,
            out List<string> sentenceCodes,
            out Dictionary<string, object> paraFormat,
            out List<string> warnings)
        {
            mode = "explicit";
            targetParagraphCodes = new List<string>();
            sentenceCodes = new List<string>();
            paraFormat = null;
            warnings = new List<string>();

            ResolveTargetParagraphCodes(args, out targetParagraphCodes, warnings);
            if (targetParagraphCodes.Count == 0)
            {
                string sentenceCsv = args != null && args.ContainsKey("target_codes")
                    ? args["target_codes"]?.ToString()?.Trim()
                    : "";
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (string code in FormatContextHelper.ParseSentenceCodes(sentenceCsv))
                {
                    if (seen.Add(code))
                    {
                        sentenceCodes.Add(code);
                    }
                }
            }

            if (targetParagraphCodes.Count == 0 && sentenceCodes.Count == 0)
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
