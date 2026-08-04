using System;
using System.Collections.Generic;
using System.Linq;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    public sealed class DocumentActionAmbiguityAnchor
    {
        public int ActionIndex { get; set; }
        public string Field { get; set; }
        public string Codes { get; set; }
        public string TableId { get; set; }

        public string InsertContentPlaceholder { get; set; }

        [Obsolete("批次 1 起废止 index，保留字段仅供未改工具编译")]
        public bool HasIndexField { get; set; }

        [Obsolete("批次 1 起废止 indexes，保留字段仅供未改工具编译")]
        public bool HasIndexesField { get; set; }

        [Obsolete("批次 1 起废止 index，保留字段仅供未改工具编译")]
        public int? Index { get; set; }

        [Obsolete("批次 1 起废止 indexes，保留字段仅供未改工具编译")]
        public string Indexes { get; set; }
    }

    public sealed class AmbiguityHintContext
    {
        public string ToolName { get; set; } = "F_process_document_actions";

        public Func<DocumentActionAmbiguityAnchor, string> FormatAnchorLabel { get; set; }

        [Obsolete("P_/apply 批次 2/3 前仍引用；S_ 改字 hint 不再使用 index 参数名")]
        public string IndexParameterName { get; set; } = "index";

        public AmbiguityActionKind ActionKind { get; set; } = AmbiguityActionKind.Other;

        public string InsertContentPlaceholder { get; set; }
    }

    public sealed class AmbiguityCheckResult
    {
        public bool HasBlockingError { get; set; }
        public string ErrorCode { get; set; }
        public string SummaryMessage { get; set; }
        public string Hint { get; set; }
        public List<Dictionary<string, object>> AmbiguousCodes { get; set; } = new List<Dictionary<string, object>>();

        public Dictionary<string, int> ResolvedDisplayStartPositionByField { get; set; } =
            new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>P_ 路径（批次 3 前仍用 index）— S_ 改字/套用请用 ResolvedDisplayStartPositionByField。</summary>
        public Dictionary<string, List<int>> ResolvedOccurrenceIndexesByField { get; set; } =
            new Dictionary<string, List<int>>(StringComparer.Ordinal);
    }

    public static class SentenceCodeAmbiguityHelper
    {
        public const string ErrorAmbiguousSentenceCode = DisplaySequenceLocator.ErrorAmbiguousSentenceCode;
        public const string ErrorInvalidAnchorSequence = DisplaySequenceLocator.ErrorInvalidAnchorSequence;
        public const string ErrorUnknownTableId = DisplaySequenceLocator.ErrorUnknownTableId;
        public const string ErrorSentenceNotInTable = DisplaySequenceLocator.ErrorSentenceNotInTable;

        public const int ContextTextPreviewMaxChars = LlmFailureExplanationBuilder.ContextTextPreviewMaxChars;

        public static string MakeFieldKey(int actionIndex, string field)
        {
            return actionIndex + ":" + (field ?? "");
        }

        public static AmbiguityCheckResult CheckBatch(
            IReadOnlyList<DocumentActionAmbiguityAnchor> anchors,
            string snapshot,
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
                    .Where(c => c.StartsWith("S_", StringComparison.Ordinal))
                    .ToList();

                if (codeList.Count == 0)
                {
                    continue;
                }

                var options = new DisplaySequenceLocateOptions
                {
                    TableId = anchor.TableId,
                    Snapshot = snapshot,
                    SegSnapshot = DocumentState.SegSnapshot,
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
                result.ErrorCode = locateResult.ErrorCode ?? ErrorInvalidAnchorSequence;

                if (string.Equals(result.ErrorCode, ErrorAmbiguousSentenceCode, StringComparison.Ordinal))
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
                && string.Equals(result.ErrorCode, ErrorAmbiguousSentenceCode, StringComparison.Ordinal))
            {
                result.AmbiguousCodes = ambiguousByCode.Values.ToList();
            }

            return result;
        }

        public static List<Dictionary<string, object>> BuildAmbiguousEntries(
            DocumentActionAmbiguityAnchor anchor,
            List<string> ambiguousCodeList,
            string snapshot,
            AppConfig.DocumentActionsAmbiguitySettings contextSettings,
            TableScopeIndex tableScope = null)
        {
            List<string> codeList = ambiguousCodeList ?? ParseCodeList(anchor?.Codes);
            var options = new DisplaySequenceLocateOptions
            {
                TableId = anchor?.TableId,
                Snapshot = snapshot,
                SegSnapshot = DocumentState.SegSnapshot,
                TableScope = tableScope
            };

            DisplaySequenceLocateResult locateResult = DisplaySequenceLocator.TryResolve(codeList, options);
            return LlmFailureExplanationBuilder.BuildAmbiguousEntries(
                anchor,
                codeList,
                locateResult,
                contextSettings);
        }

        public static Dictionary<string, object> BuildOccurrenceContext(
            IReadOnlyList<string> snapshotCodes,
            int snapshotPosition,
            string targetCode,
            AppConfig.DocumentActionsAmbiguitySettings settings)
        {
            int[] segIndices = DocumentState.BuildSegIndexBySnapshotPosition();
            return LlmFailureExplanationBuilder.BuildOccurrenceContext(
                snapshotCodes,
                snapshotPosition,
                targetCode,
                settings,
                segIndices);
        }

        public static ToolResult BuildPreStep3FailureResult(
            string documentName,
            int totalActions,
            AmbiguityCheckResult ambiguity)
        {
            var data = new Dictionary<string, object>
            {
                ["error"] = ambiguity.ErrorCode ?? "",
                ["document_name"] = documentName ?? "",
                ["total_actions"] = totalActions,
                ["executed_actions"] = 0,
                ["checkpoint_cleanup_only"] = true
            };

            if (!string.IsNullOrEmpty(ambiguity.Hint))
            {
                data["hint"] = ambiguity.Hint;
            }

            if (ambiguity.AmbiguousCodes != null && ambiguity.AmbiguousCodes.Count > 0)
            {
                data["ambiguous_codes"] = ambiguity.AmbiguousCodes;
            }

            return new ToolResult
            {
                Success = false,
                Error = ambiguity.SummaryMessage ?? "操作失败",
                Data = data
            };
        }

        public static ToolResult BuildPreStep3FailureResult(
            string documentName,
            int totalActions,
            string errorCode,
            string message)
        {
            return new ToolResult
            {
                Success = false,
                Error = message ?? "操作失败",
                Data = new Dictionary<string, object>
                {
                    ["error"] = errorCode ?? "",
                    ["document_name"] = documentName ?? "",
                    ["total_actions"] = totalActions,
                    ["executed_actions"] = 0,
                    ["checkpoint_cleanup_only"] = true
                }
            };
        }

        public static ToolResult BuildApplyAmbiguityFailureResult(
            string toolName,
            AmbiguityCheckResult ambiguity)
        {
            var data = new Dictionary<string, object>
            {
                ["error"] = ambiguity?.ErrorCode ?? "",
                ["checkpoint_cleanup_only"] = true,
                ["success"] = false
            };

            if (!string.IsNullOrEmpty(ambiguity?.Hint))
            {
                data["hint"] = ambiguity.Hint;
            }

            if (ambiguity?.AmbiguousCodes != null && ambiguity.AmbiguousCodes.Count > 0)
            {
                data["ambiguous_codes"] = ambiguity.AmbiguousCodes;
            }

            return new ToolResult
            {
                Success = false,
                Error = ambiguity?.SummaryMessage ?? "格式套用失败",
                Data = data
            };
        }

        private static AmbiguityHintProfile BuildHintProfile(
            DocumentActionAmbiguityAnchor anchor,
            AmbiguityHintContext hintContext)
        {
            AmbiguityActionKind actionKind = hintContext?.ActionKind ?? AmbiguityActionKind.Other;
            if (actionKind == AmbiguityActionKind.Other)
            {
                actionKind = LlmFailureExplanationBuilder.ResolveActionKind(anchor?.Field);
            }

            return new AmbiguityHintProfile
            {
                ToolName = string.IsNullOrEmpty(hintContext?.ToolName)
                    ? "F_process_document_actions"
                    : hintContext.ToolName,
                Field = anchor?.Field,
                ActionKind = actionKind,
                InsertContentPlaceholder = hintContext?.InsertContentPlaceholder
                    ?? anchor?.InsertContentPlaceholder,
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

            if (string.Equals(errorCode, ErrorSentenceNotInTable, StringComparison.Ordinal))
            {
                return $"编码不在表格 {anchor?.TableId} 内，无法定位 {label}（codes={codes}）";
            }

            return $"锚点 codes 无效或不连续/跨 seg，无法定位 {label}（codes={codes}）";
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
