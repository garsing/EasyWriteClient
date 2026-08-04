using System;
using System.Collections.Generic;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    /// <summary>S_ 套格式 / 读 context 的表域参数（无 index）。</summary>
    public sealed class ApplyFormatScopeArgs
    {
        public string TableId;
        public string InheritFromTableId;
    }

    public static class ApplyFormatAmbiguityHelper
    {
        public static ApplyFormatScopeArgs ParseScopeArgs(Dictionary<string, object> args)
        {
            var result = new ApplyFormatScopeArgs();
            if (args == null)
            {
                return result;
            }

            if (args.TryGetValue("in_table", out object tableObj) && tableObj != null)
            {
                result.TableId = tableObj.ToString()?.Trim();
            }

            if (args.TryGetValue("inherit_from_table_id", out object inheritTableObj) && inheritTableObj != null)
            {
                result.InheritFromTableId = inheritTableObj.ToString()?.Trim();
            }

            return result;
        }

        public static TableScopeIndex BuildTableScopeOrNull(ApplyFormatScopeArgs scopeArgs, out string cacheError)
        {
            cacheError = null;
            bool needsTableScope = !string.IsNullOrEmpty(scopeArgs?.TableId)
                || !string.IsNullOrEmpty(scopeArgs?.InheritFromTableId);
            if (!needsTableScope)
            {
                return null;
            }

            if (DocumentState.LastProcessCache == null)
            {
                cacheError = "请先调用 F_get_document_content 刷新文档映射后再使用 in_table";
                return null;
            }

            return TableSentenceScopeHelper.BuildIndexFromProcessCache(DocumentState.LastProcessCache);
        }

        public static List<DocumentActionAmbiguityAnchor> BuildAnchors(
            string mode,
            string inheritFrom,
            IReadOnlyList<string> targetCodes,
            ApplyFormatScopeArgs scopeArgs)
        {
            var anchors = new List<DocumentActionAmbiguityAnchor>();
            scopeArgs = scopeArgs ?? new ApplyFormatScopeArgs();
            targetCodes = targetCodes ?? new List<string>();

            if (string.Equals(mode, "inherit", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrEmpty(inheritFrom))
            {
                anchors.Add(new DocumentActionAmbiguityAnchor
                {
                    ActionIndex = 0,
                    Field = "inherit_from",
                    Codes = inheritFrom,
                    TableId = scopeArgs.InheritFromTableId
                });
            }

            var validTargetCodes = new List<string>();
            foreach (string code in targetCodes)
            {
                if (!string.IsNullOrWhiteSpace(code))
                {
                    validTargetCodes.Add(code.Trim());
                }
            }

            if (validTargetCodes.Count > 0)
            {
                anchors.Add(new DocumentActionAmbiguityAnchor
                {
                    ActionIndex = 0,
                    Field = "target_codes",
                    Codes = string.Join(",", validTargetCodes),
                    TableId = scopeArgs.TableId
                });
            }

            return anchors;
        }

        public static AmbiguityHintContext CreateApplyHintContext(string toolName)
        {
            return new AmbiguityHintContext
            {
                ToolName = toolName,
                ActionKind = AmbiguityActionKind.FormatApply
            };
        }

        public static AmbiguityHintContext CreateContextReadHintContext(string toolName)
        {
            return new AmbiguityHintContext
            {
                ToolName = toolName,
                ActionKind = AmbiguityActionKind.FormatContextRead
            };
        }

        public static AmbiguityCheckResult RunPrecheck(
            string mode,
            string inheritFrom,
            IReadOnlyList<string> targetCodes,
            ApplyFormatScopeArgs scopeArgs,
            string toolName,
            out TableScopeIndex tableScope,
            out string cacheError)
        {
            return RunPrecheckInternal(
                mode,
                inheritFrom,
                targetCodes,
                scopeArgs,
                toolName,
                CreateApplyHintContext(toolName),
                out tableScope,
                out cacheError);
        }

        public static AmbiguityCheckResult RunContextReadPrecheck(
            IReadOnlyList<string> targetCodes,
            ApplyFormatScopeArgs scopeArgs,
            string toolName,
            out TableScopeIndex tableScope,
            out string cacheError)
        {
            return RunPrecheckInternal(
                "",
                "",
                targetCodes,
                scopeArgs,
                toolName,
                CreateContextReadHintContext(toolName),
                out tableScope,
                out cacheError);
        }

        private static AmbiguityCheckResult RunPrecheckInternal(
            string mode,
            string inheritFrom,
            IReadOnlyList<string> targetCodes,
            ApplyFormatScopeArgs scopeArgs,
            string toolName,
            AmbiguityHintContext hintContext,
            out TableScopeIndex tableScope,
            out string cacheError)
        {
            tableScope = BuildTableScopeOrNull(scopeArgs, out cacheError);
            if (cacheError != null)
            {
                return null;
            }

            var anchors = BuildAnchors(mode, inheritFrom, targetCodes, scopeArgs);
            if (anchors.Count == 0)
            {
                return new AmbiguityCheckResult();
            }

            return SentenceCodeAmbiguityHelper.CheckBatch(
                anchors,
                DocumentState.Snapshot,
                ConfigManager.GetDocumentActionsAmbiguitySettings(),
                tableScope,
                hintContext);
        }

        public static int GetDisplayStartAt(
            AmbiguityCheckResult ambiguityResult,
            int actionIndex,
            string field)
        {
            string key = SentenceCodeAmbiguityHelper.MakeFieldKey(actionIndex, field);
            if (ambiguityResult?.ResolvedDisplayStartPositionByField != null
                && ambiguityResult.ResolvedDisplayStartPositionByField.TryGetValue(key, out int displayStart))
            {
                return displayStart;
            }

            return -1;
        }
    }
}
