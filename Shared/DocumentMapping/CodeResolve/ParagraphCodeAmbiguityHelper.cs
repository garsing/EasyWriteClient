using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// P_ 段落 display 有序流工具 + Locator 批检（无 index）。
    /// </summary>
    public static class ParagraphCodeAmbiguityHelper
    {
        public const string ErrorUnknownTableId = DisplaySequenceLocator.ErrorUnknownTableId;
        public const string ErrorParagraphNotInTable = DisplaySequenceLocator.ErrorSentenceNotInTable;

        /// <summary>与 S_ 对齐，统一 ambiguous_sentence_code。</summary>
        public const string ErrorAmbiguousParagraphCode = DisplaySequenceLocator.ErrorAmbiguousSentenceCode;

        private static readonly Regex ParagraphStartRegex = new Regex(
            @"\[((P_[0-9a-zA-Z]{5}))-start\]",
            RegexOptions.Compiled);

        public const int ContextTextPreviewMaxChars = LlmFailureExplanationBuilder.ContextTextPreviewMaxChars;

        public static bool IsEmptyParagraphStoredText(string storedText)
        {
            if (string.IsNullOrEmpty(storedText))
            {
                return true;
            }

            return storedText.Trim('\r', '\n', ' ', '\t') == string.Empty;
        }

        public static List<string> GetOrderedParagraphCodesFromDisplay()
        {
            var codes = new List<string>();
            if (DocumentState.ChunkParagraphDisplayContents == null)
            {
                return codes;
            }

            string combined = string.Concat(DocumentState.ChunkParagraphDisplayContents);
            foreach (Match match in ParagraphStartRegex.Matches(combined))
            {
                codes.Add(match.Groups[1].Value);
            }

            return codes;
        }

        public static List<int> GetDisplayPositionsForParagraphCode(string paragraphCode)
        {
            var positions = new List<int>();
            if (string.IsNullOrEmpty(paragraphCode))
            {
                return positions;
            }

            List<string> ordered = GetOrderedParagraphCodesFromDisplay();
            for (int i = 0; i < ordered.Count; i++)
            {
                if (string.Equals(ordered[i], paragraphCode, StringComparison.Ordinal))
                {
                    positions.Add(i);
                }
            }

            return positions;
        }

        public static int CountParagraphCodeInDisplay(string paragraphCode)
        {
            return GetDisplayPositionsForParagraphCode(paragraphCode).Count;
        }

        public static AmbiguityCheckResult CheckBatch(
            IReadOnlyList<DocumentActionAmbiguityAnchor> anchors,
            AppConfig.DocumentActionsAmbiguitySettings contextSettings,
            TableScopeIndex tableScope = null,
            AmbiguityHintContext hintContext = null)
        {
            var result = new AmbiguityCheckResult();
            if (anchors == null || anchors.Count == 0)
            {
                return result;
            }

            var ambiguousByCode = new Dictionary<string, Dictionary<string, object>>(StringComparer.Ordinal);

            foreach (DocumentActionAmbiguityAnchor anchor in anchors)
            {
                if (string.IsNullOrWhiteSpace(anchor?.Codes))
                {
                    continue;
                }

                List<string> codeList = ParseCodeList(anchor.Codes)
                    .Where(c => c.StartsWith("P_", StringComparison.Ordinal))
                    .ToList();

                if (codeList.Count == 0)
                {
                    continue;
                }

                var options = new DisplaySequenceLocateOptions
                {
                    StreamKind = CodeStreamKind.ParagraphDisplay,
                    TableId = anchor.TableId,
                    TableScope = tableScope
                };

                DisplaySequenceLocateResult locateResult = DisplaySequenceLocator.TryResolve(codeList, options);
                if (locateResult.MatchCount == 1 && locateResult.DisplayStartPosition.HasValue)
                {
                    result.ResolvedDisplayStartPositionByField[MakeFieldKey(anchor.ActionIndex, anchor.Field)] =
                        locateResult.DisplayStartPosition.Value;
                    continue;
                }

                result.HasBlockingError = true;
                result.ErrorCode = locateResult.ErrorCode ?? DisplaySequenceLocator.ErrorInvalidAnchorSequence;

                if (string.Equals(result.ErrorCode, ErrorAmbiguousParagraphCode, StringComparison.Ordinal))
                {
                    AmbiguityHintProfile profile = BuildHintProfile(anchor, hintContext);
                    List<Dictionary<string, object>> ambiguousEntries = LlmFailureExplanationBuilder.BuildAmbiguousEntries(
                        anchor,
                        codeList,
                        locateResult,
                        contextSettings,
                        profile);

                    result.SummaryMessage = LlmFailureExplanationBuilder.BuildSummaryMessage(
                        anchor,
                        codeList,
                        locateResult,
                        profile);
                    result.Hint = LlmFailureExplanationBuilder.BuildHint(
                        anchor,
                        codeList,
                        locateResult,
                        ambiguousEntries,
                        profile);

                    foreach (Dictionary<string, object> entry in ambiguousEntries)
                    {
                        string code = entry.ContainsKey("code") ? entry["code"]?.ToString() : null;
                        if (!string.IsNullOrEmpty(code) && !ambiguousByCode.ContainsKey(code))
                        {
                            ambiguousByCode[code] = entry;
                        }
                    }
                }
                else
                {
                    result.SummaryMessage = BuildInvalidAnchorMessage(anchor, codeList, locateResult.ErrorCode);
                }
            }

            if (result.HasBlockingError
                && string.Equals(result.ErrorCode, ErrorAmbiguousParagraphCode, StringComparison.Ordinal))
            {
                result.AmbiguousCodes = ambiguousByCode.Values.ToList();
            }

            return result;
        }

        public static Dictionary<string, object> BuildOccurrenceContext(
            IReadOnlyList<string> orderedCodes,
            int displayPosition,
            string targetCode,
            AppConfig.DocumentActionsAmbiguitySettings settings)
        {
            settings = NormalizeContextSettings(settings);
            var left = new List<Dictionary<string, object>>();
            var right = new List<Dictionary<string, object>>();

            for (int d = 1; d <= settings.ContextLeftSentences && displayPosition - d >= 0; d++)
            {
                left.Add(BuildParagraphPreview(orderedCodes[displayPosition - d]));
            }

            for (int d = 1; d <= settings.ContextRightSentences && displayPosition + d < orderedCodes.Count; d++)
            {
                right.Add(BuildParagraphPreview(orderedCodes[displayPosition + d]));
            }

            return new Dictionary<string, object>
            {
                ["left"] = left,
                ["target"] = BuildParagraphPreview(targetCode),
                ["right"] = right,
            };
        }

        public static string MakeFieldKey(int actionIndex, string field)
        {
            return actionIndex + ":" + (field ?? "");
        }

        private static AmbiguityHintProfile BuildHintProfile(
            DocumentActionAmbiguityAnchor anchor,
            AmbiguityHintContext hintContext)
        {
            AmbiguityActionKind actionKind = hintContext?.ActionKind ?? AmbiguityActionKind.ParagraphFormatApply;
            if (actionKind == AmbiguityActionKind.Other)
            {
                actionKind = AmbiguityActionKind.ParagraphFormatApply;
            }

            return new AmbiguityHintProfile
            {
                ToolName = string.IsNullOrEmpty(hintContext?.ToolName)
                    ? "F_apply_paragraph_format"
                    : hintContext.ToolName,
                Field = anchor?.Field,
                ActionKind = actionKind,
            };
        }

        private static string BuildInvalidAnchorMessage(
            DocumentActionAmbiguityAnchor anchor,
            IReadOnlyList<string> codeList,
            string errorCode)
        {
            string label = $"action[{anchor?.ActionIndex ?? 0}].{anchor?.Field ?? "codes"}";
            string codes = string.Join(",", codeList ?? Array.Empty<string>());

            if (string.Equals(errorCode, ErrorUnknownTableId, StringComparison.Ordinal))
            {
                return $"表格 {anchor?.TableId} 不存在，无法定位 {label}（codes={codes}）";
            }

            if (string.Equals(errorCode, ErrorParagraphNotInTable, StringComparison.Ordinal))
            {
                return $"编码不在表格 {anchor?.TableId} 内，无法定位 {label}（codes={codes}）";
            }

            return $"锚点 codes 无效或不连续，无法定位 {label}（codes={codes}）";
        }

        private static Dictionary<string, object> BuildParagraphPreview(string code)
        {
            string text = DocumentState.GetParagraphContent(code) ?? "";
            return new Dictionary<string, object>
            {
                ["code"] = code,
                ["text"] = TruncatePreview(text),
            };
        }

        private static string TruncatePreview(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            if (text.Length <= ContextTextPreviewMaxChars)
            {
                return text;
            }

            return text.Substring(0, ContextTextPreviewMaxChars) + "...";
        }

        private static AppConfig.DocumentActionsAmbiguitySettings NormalizeContextSettings(
            AppConfig.DocumentActionsAmbiguitySettings settings)
        {
            settings = settings ?? new AppConfig.DocumentActionsAmbiguitySettings();
            return new AppConfig.DocumentActionsAmbiguitySettings
            {
                ContextLeftSentences = Math.Max(0, Math.Min(5, settings.ContextLeftSentences)),
                ContextRightSentences = Math.Max(0, Math.Min(5, settings.ContextRightSentences)),
            };
        }

        private static List<string> ParseCodeList(string codes)
        {
            if (string.IsNullOrWhiteSpace(codes))
            {
                return new List<string>();
            }

            return codes.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }
    }
}
