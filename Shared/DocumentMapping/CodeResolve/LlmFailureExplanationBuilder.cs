using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordAddIn1;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    public enum AmbiguityActionKind
    {
        Replace = 0,
        Delete = 1,
        InsertTarget = 2,
        Other = 3,
        FormatApply = 4,
        FormatContextRead = 5,
        ParagraphFormatApply = 6,
        ParagraphContextRead = 7
    }

    public sealed class AmbiguityHintProfile
    {
        public string ToolName { get; set; } = "F_process_document_actions";

        public string Field { get; set; }

        public AmbiguityActionKind ActionKind { get; set; } = AmbiguityActionKind.Other;

        public string InsertContentPlaceholder { get; set; }
    }

    public static class LlmFailureExplanationBuilder
    {
        public const int ContextTextPreviewMaxChars = 120;

        public static AmbiguityActionKind ResolveActionKind(string field)
        {
            if (string.Equals(field, "original", StringComparison.Ordinal)
                || string.Equals(field, "original.codes", StringComparison.Ordinal))
            {
                return AmbiguityActionKind.Replace;
            }

            if (string.Equals(field, "deletion", StringComparison.Ordinal)
                || string.Equals(field, "deletion.codes", StringComparison.Ordinal))
            {
                return AmbiguityActionKind.Delete;
            }

            if (string.Equals(field, "target", StringComparison.Ordinal)
                || string.Equals(field, "target.codes", StringComparison.Ordinal))
            {
                return AmbiguityActionKind.InsertTarget;
            }

            return AmbiguityActionKind.Other;
        }

        public static List<Dictionary<string, object>> BuildAmbiguousEntries(
            DocumentActionAmbiguityAnchor anchor,
            IReadOnlyList<string> anchorCodeList,
            DisplaySequenceLocateResult locateResult,
            AppConfig.DocumentActionsAmbiguitySettings contextSettings,
            AmbiguityHintProfile profile = null)
        {
            profile = profile ?? new AmbiguityHintProfile
            {
                Field = anchor?.Field,
                ActionKind = ResolveActionKind(anchor?.Field)
            };

            var entries = new List<Dictionary<string, object>>();
            if (anchorCodeList == null || anchorCodeList.Count == 0 || locateResult?.MatchStartPositions == null)
            {
                return entries;
            }

            IReadOnlyList<string> domain = locateResult.OrderedDomainCodes ?? Array.Empty<string>();
            int[] segIndices = locateResult.SegIndexByPosition;
            var settings = NormalizeContextSettings(contextSettings);
            string primaryCode = anchorCodeList[0];
            var occurrences = new List<Dictionary<string, object>>();

            foreach (int startPos in locateResult.MatchStartPositions)
            {
                string targetCode = startPos >= 0 && startPos < domain.Count ? domain[startPos] : primaryCode;
                var context = BuildOccurrenceContext(domain, startPos, targetCode, settings, segIndices);
                string followingContextText = BuildFollowingContextText(context);
                string suggestedLocatorCodes = DisplaySequenceLocator.TryExpandSuggestedLocatorCodes(
                    anchorCodeList,
                    startPos,
                    domain,
                    segIndices);

                var occurrence = new Dictionary<string, object>
                {
                    ["display_start_position"] = startPos,
                    ["context"] = context,
                    ["following_context_text"] = followingContextText,
                    ["suggested_locator_codes"] = suggestedLocatorCodes ?? (object)null
                };

                AppendStructuredNewSuggestions(
                    occurrence,
                    profile.ActionKind,
                    context,
                    targetCode,
                    profile.InsertContentPlaceholder ?? anchor?.InsertContentPlaceholder);

                occurrences.Add(occurrence);
            }

            entries.Add(new Dictionary<string, object>
            {
                ["code"] = primaryCode,
                ["occurrence_count"] = locateResult.MatchStartPositions.Count,
                ["action_index"] = anchor.ActionIndex,
                ["field"] = anchor.Field,
                ["in_table"] = string.IsNullOrEmpty(anchor.TableId) ? null : anchor.TableId,
                ["occurrences"] = occurrences
            });

            return entries;
        }

        public static string BuildSummaryMessage(
            DocumentActionAmbiguityAnchor anchor,
            IReadOnlyList<string> anchorCodeList,
            DisplaySequenceLocateResult locateResult,
            AmbiguityHintProfile profile = null)
        {
            profile = profile ?? new AmbiguityHintProfile { Field = anchor?.Field };
            string label = FormatAnchorLabel(anchor, profile);
            string code = anchorCodeList?.FirstOrDefault() ?? "";
            int count = locateResult?.MatchStartPositions?.Count ?? 0;
            string domainLabel = string.IsNullOrEmpty(anchor?.TableId)
                ? "全文"
                : $"表格 {anchor.TableId}";

            return
                $"编码 {code} 在{domainLabel}中有 {count} 处匹配，无法唯一定位，请为 {label} 加上下文后重试";
        }

        public static string BuildHint(
            DocumentActionAmbiguityAnchor anchor,
            IReadOnlyList<string> anchorCodeList,
            DisplaySequenceLocateResult locateResult,
            List<Dictionary<string, object>> ambiguousEntries,
            AmbiguityHintProfile profile = null)
        {
            profile = profile ?? new AmbiguityHintProfile
            {
                Field = anchor?.Field,
                ActionKind = ResolveActionKind(anchor?.Field)
            };

            if (profile.ActionKind == AmbiguityActionKind.FormatApply
                || profile.ActionKind == AmbiguityActionKind.FormatContextRead
                || profile.ActionKind == AmbiguityActionKind.ParagraphFormatApply
                || profile.ActionKind == AmbiguityActionKind.ParagraphContextRead)
            {
                return BuildFormatDisambiguationHint(
                    anchor,
                    anchorCodeList,
                    locateResult,
                    ambiguousEntries,
                    profile);
            }

            string code = anchorCodeList?.FirstOrDefault() ?? "";
            int count = locateResult?.MatchStartPositions?.Count ?? 0;
            string domainLabel = string.IsNullOrEmpty(anchor?.TableId) ? "全文" : "表内";
            string toolName = string.IsNullOrEmpty(profile.ToolName)
                ? "F_process_document_actions"
                : profile.ToolName;

            var builder = new StringBuilder();
            builder.AppendLine(
                $"你提供的编码 {code} 在{domainLabel}中有 {count} 处匹配，无法唯一定位，本次调用已终止，文档未被修改。相同句子文本映射为同一 {code}。");
            builder.AppendLine();
            builder.AppendLine(
                "请加上下文（后续句子）实现更精确的定位：重试时将下方「定位 codes」整串写入入参 codes（逗号分隔，无 after_codes 字段）。"
                + "codes 须 display 中严格相邻且同一 seg；**不要求每个码都全文唯一**，只要整串子串恰好匹配 1 次即可（重复码可与唯一邻句同串）。");

            Dictionary<string, object> entry = ambiguousEntries?.FirstOrDefault();
            var occurrences = entry?["occurrences"] as List<Dictionary<string, object>>;
            if (occurrences != null)
            {
                for (int i = 0; i < occurrences.Count; i++)
                {
                    string following = occurrences[i].ContainsKey("following_context_text")
                        ? occurrences[i]["following_context_text"]?.ToString()
                        : "";
                    string suggested = occurrences[i].ContainsKey("suggested_locator_codes")
                        ? occurrences[i]["suggested_locator_codes"]?.ToString()
                        : null;

                    builder.AppendLine();
                    if (string.IsNullOrEmpty(suggested))
                    {
                        builder.AppendLine(
                            $"第 {i + 1} 处：下文是「{following}」→ 下文不足以唯一定位，请改用全文唯一的邻句或 in_table");
                    }
                    else
                    {
                        builder.AppendLine(
                            $"第 {i + 1} 处：下文是「{following}」→ 定位 codes 为 {suggested}");
                        AppendPerOccurrenceActionHint(builder, profile.ActionKind, occurrences[i], i + 1);
                    }
                }
            }

            builder.AppendLine();
            builder.AppendLine(BuildActionSummary(profile.ActionKind));
            builder.AppendLine($"请重新调用 {toolName}。");
            builder.AppendLine("禁止：index、默认第一次匹配、after_codes 独立字段。");

            return builder.ToString().TrimEnd();
        }

        public static Dictionary<string, object> BuildOccurrenceContext(
            IReadOnlyList<string> domainCodes,
            int displayPosition,
            string targetCode,
            AppConfig.DocumentActionsAmbiguitySettings settings,
            int[] segIndices = null)
        {
            settings = NormalizeContextSettings(settings);
            var left = new List<Dictionary<string, object>>();
            var right = new List<Dictionary<string, object>>();
            int seg = GetSegIndex(segIndices, displayPosition);

            for (int d = 1; d <= settings.ContextLeftSentences && displayPosition - d >= 0; d++)
            {
                int pos = displayPosition - d;
                if (GetSegIndex(segIndices, pos) != seg)
                {
                    break;
                }

                left.Add(BuildSentencePreview(domainCodes[pos]));
            }

            for (int d = 1; d <= settings.ContextRightSentences && displayPosition + d < domainCodes.Count; d++)
            {
                int pos = displayPosition + d;
                if (GetSegIndex(segIndices, pos) != seg)
                {
                    break;
                }

                right.Add(BuildSentencePreview(domainCodes[pos]));
            }

            return new Dictionary<string, object>
            {
                ["left"] = left,
                ["target"] = BuildSentencePreview(targetCode),
                ["right"] = right
            };
        }

        private static void AppendStructuredNewSuggestions(
            Dictionary<string, object> occurrence,
            AmbiguityActionKind actionKind,
            Dictionary<string, object> context,
            string targetCode,
            string insertPlaceholder)
        {
            if (actionKind == AmbiguityActionKind.Delete)
            {
                occurrence["suggested_replace_new_for_delete"] = BuildFollowingTextFromContext(context);
            }
            else if (actionKind == AmbiguityActionKind.InsertTarget)
            {
                string targetText = GetPreviewText(context, "target");
                string followingText = BuildFollowingTextFromContext(context);
                string insertText = string.IsNullOrEmpty(insertPlaceholder) ? "（insert 内容）" : insertPlaceholder;
                // 仅后插：改 replace 时 new = 目标句 + insert + 下文
                occurrence["suggested_replace_new_for_insert"] = targetText + insertText + followingText;
            }
        }

        private static void AppendPerOccurrenceActionHint(
            StringBuilder builder,
            AmbiguityActionKind actionKind,
            Dictionary<string, object> occurrence,
            int occurrenceNumber)
        {
            switch (actionKind)
            {
                case AmbiguityActionKind.Replace:
                    builder.AppendLine(
                        $"  replace 重试：new = 修改后的目标句 + 保留下文「{occurrence["following_context_text"]}」");
                    break;
                case AmbiguityActionKind.Delete:
                    builder.AppendLine(
                        $"  若原意 delete：改 replace，new 仅保留下文「{occurrence.GetValueOrDefault("suggested_replace_new_for_delete")}」，不含目标句");
                    break;
                case AmbiguityActionKind.InsertTarget:
                    builder.AppendLine(
                        $"  若原意 target insert：改 replace，new = 目标句 + insert + 下文（见 suggested_replace_new_for_insert）");
                    break;
            }
        }

        private static string BuildActionSummary(AmbiguityActionKind actionKind)
        {
            switch (actionKind)
            {
                case AmbiguityActionKind.Replace:
                    return "replace 重试：new = 改目标句 + 下文原样。";
                case AmbiguityActionKind.Delete:
                    return "delete 歧义重试：改 replace，new = 仅下文（去掉目标句）。";
                case AmbiguityActionKind.InsertTarget:
                    return "insert 歧义重试：改 replace，new = 目标句 + insert + 下文。";
                case AmbiguityActionKind.FormatApply:
                    return "重试时将 suggested_locator_codes 写入 target_codes 或 inherit_from；"
                           + "多个 target 均会被套用格式。";
                case AmbiguityActionKind.FormatContextRead:
                    return "重试时将 suggested_locator_codes 写入 target_codes；"
                           + "多个 target 均会被读取。";
                case AmbiguityActionKind.ParagraphFormatApply:
                    return "重试时将 suggested_locator_codes 写入 target_paragraph_codes 或 inherit_from_paragraph；"
                           + "多个 target 均会被套用段落格式。";
                case AmbiguityActionKind.ParagraphContextRead:
                    return "重试时将 suggested_locator_codes 写入 target_paragraph_codes；"
                           + "多个 target 均会被读取。";
                default:
                    return "重试时将 suggested_locator_codes 写入 codes。";
            }
        }

        private static string BuildFormatDisambiguationHint(
            DocumentActionAmbiguityAnchor anchor,
            IReadOnlyList<string> anchorCodeList,
            DisplaySequenceLocateResult locateResult,
            List<Dictionary<string, object>> ambiguousEntries,
            AmbiguityHintProfile profile)
        {
            string code = anchorCodeList?.FirstOrDefault() ?? "";
            int count = locateResult?.MatchStartPositions?.Count ?? 0;
            string domainLabel = string.IsNullOrEmpty(anchor?.TableId) ? "全文" : "表内";
            string toolName = string.IsNullOrEmpty(profile.ToolName)
                ? "F_apply_document_format"
                : profile.ToolName;
            string fieldLabel = ResolveFormatFieldLabel(anchor?.Field, profile.ActionKind);
            string actionLabel = profile.ActionKind == AmbiguityActionKind.FormatContextRead
                || profile.ActionKind == AmbiguityActionKind.ParagraphContextRead
                ? "读取格式上下文"
                : "套用格式";

            var builder = new StringBuilder();
            builder.AppendLine(
                $"你提供的编码 {code} 在{domainLabel}中有 {count} 处匹配，无法唯一定位，本次{actionLabel}已终止，文档未被修改。相同句子文本映射为同一 {code}。");
            builder.AppendLine();
            builder.AppendLine(
                $"请加上下文（后续句子）实现更精确的定位：重试时将下方「定位 codes」整串写入 {fieldLabel}（逗号分隔，无 after_codes 字段）。"
                + " codes 须 display 中严格相邻且同一 seg；不要求每个码都全文唯一，只要整串子串恰好匹配 1 次即可。");
            if (string.Equals(fieldLabel, "target_codes", StringComparison.Ordinal)
                || string.Equals(fieldLabel, "target_paragraph_codes", StringComparison.Ordinal))
            {
                builder.AppendLine("多个 " + fieldLabel + " 均会被" + (profile.ActionKind == AmbiguityActionKind.FormatContextRead
                    || profile.ActionKind == AmbiguityActionKind.ParagraphContextRead ? "读取" : "套用") + "。");
            }

            Dictionary<string, object> entry = ambiguousEntries?.FirstOrDefault();
            var occurrences = entry?["occurrences"] as List<Dictionary<string, object>>;
            if (occurrences != null)
            {
                for (int i = 0; i < occurrences.Count; i++)
                {
                    string following = occurrences[i].ContainsKey("following_context_text")
                        ? occurrences[i]["following_context_text"]?.ToString()
                        : "";
                    string suggested = occurrences[i].ContainsKey("suggested_locator_codes")
                        ? occurrences[i]["suggested_locator_codes"]?.ToString()
                        : null;

                    builder.AppendLine();
                    if (string.IsNullOrEmpty(suggested))
                    {
                        builder.AppendLine(
                            $"第 {i + 1} 处：下文是「{following}」→ 下文不足以唯一定位，请改用全文唯一的邻句或 in_table");
                    }
                    else
                    {
                        builder.AppendLine(
                            $"第 {i + 1} 处：下文是「{following}」→ 定位 codes 为 {suggested}");
                    }
                }
            }

            builder.AppendLine();
            builder.AppendLine(BuildActionSummary(profile.ActionKind));
            builder.AppendLine($"请重新调用 {toolName}。");
            builder.AppendLine("禁止：index、默认第一次匹配、after_codes 独立字段。");

            return builder.ToString().TrimEnd();
        }

        private static string BuildFollowingContextText(Dictionary<string, object> context)
        {
            string text = BuildFollowingTextFromContext(context);
            return string.IsNullOrEmpty(text) ? "（无右邻句）" : text;
        }

        private static string BuildFollowingTextFromContext(Dictionary<string, object> context)
        {
            if (context == null || !context.TryGetValue("right", out object rightObj))
            {
                return "";
            }

            var right = rightObj as List<Dictionary<string, object>>;
            if (right == null || right.Count == 0)
            {
                return "";
            }

            var parts = new List<string>();
            foreach (Dictionary<string, object> item in right)
            {
                string text = item.ContainsKey("text") ? item["text"]?.ToString() : "";
                if (!string.IsNullOrEmpty(text))
                {
                    parts.Add(text);
                }
            }

            return string.Concat(parts);
        }

        private static string GetPreviewText(Dictionary<string, object> context, string key)
        {
            if (context == null || !context.TryGetValue(key, out object obj))
            {
                return "";
            }

            var preview = obj as Dictionary<string, object>;
            return preview != null && preview.ContainsKey("text") ? preview["text"]?.ToString() ?? "" : "";
        }

        private static string ResolveFormatFieldLabel(string field, AmbiguityActionKind actionKind)
        {
            if (string.Equals(field, "inherit_from", StringComparison.Ordinal))
            {
                return "inherit_from";
            }

            if (string.Equals(field, "inherit_from_paragraph", StringComparison.Ordinal))
            {
                return "inherit_from_paragraph";
            }

            if (string.Equals(field, "target_paragraph_codes", StringComparison.Ordinal)
                || actionKind == AmbiguityActionKind.ParagraphFormatApply
                || actionKind == AmbiguityActionKind.ParagraphContextRead)
            {
                return "target_paragraph_codes";
            }

            return "target_codes";
        }

        private static Dictionary<string, object> BuildSentencePreview(string code)
        {
            string text = code != null && code.StartsWith("P_", StringComparison.Ordinal)
                ? DocumentState.GetParagraphContent(code) ?? ""
                : DocumentState.GetSentenceContent(code) ?? "";
            return new Dictionary<string, object>
            {
                ["code"] = code,
                ["text"] = TruncatePreview(text)
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

        private static string FormatAnchorLabel(DocumentActionAmbiguityAnchor anchor, AmbiguityHintProfile profile)
        {
            return $"action[{anchor?.ActionIndex ?? 0}].{anchor?.Field ?? profile?.Field ?? "codes"}";
        }

        private static AppConfig.DocumentActionsAmbiguitySettings NormalizeContextSettings(
            AppConfig.DocumentActionsAmbiguitySettings settings)
        {
            settings = settings ?? new AppConfig.DocumentActionsAmbiguitySettings();
            return new AppConfig.DocumentActionsAmbiguitySettings
            {
                ContextLeftSentences = Math.Max(0, Math.Min(5, settings.ContextLeftSentences)),
                ContextRightSentences = Math.Max(0, Math.Min(5, settings.ContextRightSentences))
            };
        }

        private static int GetSegIndex(int[] segIndices, int position)
        {
            if (segIndices == null || position < 0 || position >= segIndices.Length)
            {
                return 0;
            }

            return segIndices[position];
        }
    }

    internal static class DictionaryExtensions
    {
        public static object GetValueOrDefault(this Dictionary<string, object> dict, string key)
        {
            return dict != null && dict.ContainsKey(key) ? dict[key] : null;
        }
    }
}
