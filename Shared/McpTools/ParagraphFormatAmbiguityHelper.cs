using System;
using System.Collections.Generic;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    /// <summary>
    /// P_ 段落 format apply / 读 context 消歧（镜像 ApplyFormatAmbiguityHelper）。
    /// </summary>
    public static class ParagraphFormatAmbiguityHelper
    {
        public static List<DocumentActionAmbiguityAnchor> BuildAnchors(
            string inheritFromParagraph,
            IReadOnlyList<string> targetParagraphCodes,
            ApplyFormatScopeArgs scopeArgs)
        {
            var anchors = new List<DocumentActionAmbiguityAnchor>();
            scopeArgs = scopeArgs ?? new ApplyFormatScopeArgs();
            targetParagraphCodes = targetParagraphCodes ?? new List<string>();

            if (!string.IsNullOrEmpty(inheritFromParagraph))
            {
                anchors.Add(new DocumentActionAmbiguityAnchor
                {
                    ActionIndex = 0,
                    Field = "inherit_from_paragraph",
                    Codes = inheritFromParagraph,
                    TableId = scopeArgs.InheritFromTableId
                });
            }

            for (int i = 0; i < targetParagraphCodes.Count; i++)
            {
                string code = targetParagraphCodes[i];
                if (string.IsNullOrWhiteSpace(code))
                {
                    continue;
                }

                anchors.Add(new DocumentActionAmbiguityAnchor
                {
                    ActionIndex = i,
                    Field = "target_paragraph_codes",
                    Codes = code,
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
                ActionKind = AmbiguityActionKind.ParagraphFormatApply
            };
        }

        public static AmbiguityHintContext CreateContextReadHintContext(string toolName)
        {
            return new AmbiguityHintContext
            {
                ToolName = toolName,
                ActionKind = AmbiguityActionKind.ParagraphContextRead
            };
        }

        public static AmbiguityCheckResult RunPrecheck(
            string inheritFromParagraph,
            ApplyFormatScopeArgs scopeArgs,
            IReadOnlyList<string> targetParagraphCodes,
            string toolName,
            out TableScopeIndex tableScope,
            out string cacheError)
        {
            return RunPrecheckInternal(
                inheritFromParagraph,
                scopeArgs,
                targetParagraphCodes,
                toolName,
                CreateApplyHintContext(toolName),
                out tableScope,
                out cacheError);
        }

        public static AmbiguityCheckResult RunContextReadPrecheck(
            IReadOnlyList<string> targetParagraphCodes,
            ApplyFormatScopeArgs scopeArgs,
            string toolName,
            out TableScopeIndex tableScope,
            out string cacheError)
        {
            return RunPrecheckInternal(
                "",
                scopeArgs,
                targetParagraphCodes,
                toolName,
                CreateContextReadHintContext(toolName),
                out tableScope,
                out cacheError);
        }

        private static AmbiguityCheckResult RunPrecheckInternal(
            string inheritFromParagraph,
            ApplyFormatScopeArgs scopeArgs,
            IReadOnlyList<string> targetParagraphCodes,
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

            var anchors = BuildAnchors(inheritFromParagraph, targetParagraphCodes, scopeArgs);
            if (anchors.Count == 0)
            {
                return new AmbiguityCheckResult();
            }

            return ParagraphCodeAmbiguityHelper.CheckBatch(
                anchors,
                ConfigManager.GetDocumentActionsAmbiguitySettings(),
                tableScope,
                hintContext);
        }

        public static int GetDisplayStartAt(
            AmbiguityCheckResult ambiguityResult,
            int actionIndex,
            string field)
        {
            string key = ParagraphCodeAmbiguityHelper.MakeFieldKey(actionIndex, field);
            if (ambiguityResult?.ResolvedDisplayStartPositionByField != null
                && ambiguityResult.ResolvedDisplayStartPositionByField.TryGetValue(key, out int displayStart))
            {
                return displayStart;
            }

            return -1;
        }

        private static TableScopeIndex BuildTableScopeOrNull(ApplyFormatScopeArgs scopeArgs, out string cacheError)
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

            return TableParagraphScopeHelper.BuildIndexFromProcessCache(DocumentState.LastProcessCache);
        }
    }
}
