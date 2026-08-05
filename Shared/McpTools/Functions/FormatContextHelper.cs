using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    /// <summary>
    /// Phase 2：格式上下文组装（邻居遍历、fingerprint、char/para 摘要、pre_replace 读取）。
    /// </summary>
    public static class FormatContextHelper
    {
        public const string DbgPrefix = "[格式上下文][DBG]";
        public const int DefaultTextPreviewLength = 40;
        public const int DefaultContextWindow = 3;

        private static readonly string[] CharFormatKeys =
        {
            "font_name", "font_size", "font_color", "background_color",
            "is_bold", "has_underline", "has_strikethrough"
        };

        private static readonly string[] ParaFormatKeys = { "alignment" };

        public static void DbgLog(string message)
        {
            if (!FormatInheritHelper.EnableDbg)
            {
                return;
            }

            System.Diagnostics.Debug.WriteLine($"{DbgPrefix} {message}");
        }

        public sealed class FormatContextBuildResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public List<string> TargetCodes { get; set; } = new List<string>();
            public List<string> TargetParagraphCodes { get; set; } = new List<string>();
            public int ContextWindow { get; set; }
            public List<Dictionary<string, object>> Entries { get; set; } = new List<Dictionary<string, object>>();
            public List<string> Errors { get; set; } = new List<string>();
            public bool IsParagraphMode => TargetParagraphCodes.Count > 0;
        }

        public static FormatContextBuildResult BuildContext(
            Word.Document doc,
            string targetCodesCsv,
            int? contextWindow = null)
        {
            return BuildSentenceContext(doc, targetCodesCsv, contextWindow, null, null, null);
        }

        public static FormatContextBuildResult BuildContext(
            Word.Document doc,
            string targetCodesCsv,
            int? contextWindow,
            string tableId,
            TableScopeIndex tableScope,
            AmbiguityCheckResult precheckResult)
        {
            return BuildSentenceContext(doc, targetCodesCsv, contextWindow, tableId, tableScope, precheckResult);
        }

        public static FormatContextBuildResult BuildParagraphContext(
            Word.Document doc,
            string targetParagraphCodesCsv,
            int? contextWindow = null)
        {
            return BuildParagraphContext(doc, targetParagraphCodesCsv, contextWindow, null, null, null);
        }

        public static FormatContextBuildResult BuildParagraphContext(
            Word.Document doc,
            string targetParagraphCodesCsv,
            int? contextWindow,
            string tableId,
            TableScopeIndex tableScope,
            AmbiguityCheckResult precheckResult)
        {
            var result = new FormatContextBuildResult
            {
                ContextWindow = 0
            };

            if (doc == null)
            {
                result.Success = false;
                result.Error = "Word 文档不可用";
                return result;
            }

            var targetCodes = ParseParagraphCodes(targetParagraphCodesCsv);
            if (targetCodes.Count == 0)
            {
                result.Success = false;
                result.Error = "缺少或无效的 target_paragraph_codes（须为逗号分隔的 P_ 编码）";
                return result;
            }

            foreach (string code in targetCodes)
            {
                if (!IsParagraphCode(code))
                {
                    result.Success = false;
                    result.Error = "get_format_context（P_ 路径）仅支持 P_ 段落编码";
                    return result;
                }
            }

            result.TargetParagraphCodes = targetCodes;

            if (precheckResult == null)
            {
                result.Success = false;
                result.Error = "缺少段落消歧结果，请先完成 precheck";
                return result;
            }

            int displayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                precheckResult, 0, "target_paragraph_codes");
            if (displayStart < 0)
            {
                result.Success = false;
                result.Error = "target_paragraph_codes 无法在段落快照中定位";
                return result;
            }

            Word.Range span = DisplayPositionRangeResolver.ResolveParagraphSpan(
                doc,
                targetCodes,
                displayStart,
                tableId,
                tableScope,
                debugTag: "format_context_para_span");
            if (span == null)
            {
                result.Success = false;
                result.Error = "无法定位 target_paragraph_codes 序列 Span";
                return result;
            }

            // 单码时 Span 即目标段；再 Find 会与 scope 重合触发 Word Find 怪癖
            if (targetCodes.Count == 1)
            {
                Dictionary<string, object> singleEntry = BuildParagraphEntryFromRange(
                    targetCodes[0],
                    span,
                    isTarget: true);
                if (singleEntry != null)
                {
                    result.Entries.Add(singleEntry);
                }
                else
                {
                    result.Errors.Add($"无法读取段落格式: {targetCodes[0]}");
                }
            }
            else
            {
                foreach (string code in targetCodes)
                {
                    Word.Range paragraphRange = SpanParagraphLocator.LocateParagraph(
                        doc,
                        span,
                        code,
                        $"format_context:{code}");
                    Dictionary<string, object> entry = BuildParagraphEntryFromRange(
                        code,
                        paragraphRange,
                        isTarget: true);
                    if (entry != null)
                    {
                        result.Entries.Add(entry);
                    }
                    else
                    {
                        result.Errors.Add($"无法读取段落格式: {code}");
                    }
                }
            }

            result.Success = result.Entries.Count > 0;
            if (!result.Success)
            {
                result.Error = "未能组装任何 entries";
            }

            DbgLog(
                $"BuildParagraphContext targets=[{string.Join(",", targetCodes)}] " +
                $"entries={result.Entries.Count} warnings={result.Errors.Count}");
            return result;
        }

        private static FormatContextBuildResult BuildSentenceContext(
            Word.Document doc,
            string targetCodesCsv,
            int? contextWindow,
            string tableId,
            TableScopeIndex tableScope,
            AmbiguityCheckResult precheckResult)
        {
            var result = new FormatContextBuildResult
            {
                ContextWindow = 0
            };

            if (doc == null)
            {
                result.Success = false;
                result.Error = "Word 文档不可用";
                return result;
            }

            var targetCodes = ParseSentenceCodes(targetCodesCsv);
            if (targetCodes.Count == 0)
            {
                result.Success = false;
                result.Error = "缺少或无效的 target_codes（须为逗号分隔的 S_ 编码）";
                return result;
            }

            foreach (string code in targetCodes)
            {
                if (code.StartsWith("T_", StringComparison.Ordinal))
                {
                    result.Success = false;
                    result.Error = "get_format_context 仅支持 S_ 句子编码，不支持 T_ 表格编码";
                    return result;
                }
            }

            result.TargetCodes = targetCodes;

            if (precheckResult == null)
            {
                result.Success = false;
                result.Error = "缺少句子消歧结果，请先完成 precheck";
                return result;
            }

            int displayStart = ApplyFormatAmbiguityHelper.GetDisplayStartAt(
                precheckResult, 0, "target_codes");
            if (displayStart < 0)
            {
                result.Success = false;
                result.Error = "target_codes 无法在文档快照中定位";
                return result;
            }

            Word.Range span = DisplayPositionRangeResolver.ResolveSpan(
                doc,
                targetCodes,
                displayStart,
                tableId,
                tableScope,
                debugTag: "format_context_span");
            if (span == null)
            {
                result.Success = false;
                result.Error = "无法定位 target_codes 序列 Span";
                return result;
            }

            // 单码时 Span 即目标句；再 Find 会与 scope 重合触发 Word Find 怪癖
            if (targetCodes.Count == 1)
            {
                Dictionary<string, object> singleEntry = BuildEntryFromRange(
                    doc,
                    targetCodes[0],
                    span,
                    isTarget: true);
                if (singleEntry != null)
                {
                    result.Entries.Add(singleEntry);
                }
                else
                {
                    result.Errors.Add($"无法读取格式: {targetCodes[0]}");
                }
            }
            else
            {
                foreach (string code in targetCodes)
                {
                    Word.Range range = SpanSentenceLocator.LocateSentence(
                        doc,
                        span,
                        code,
                        $"format_context:{code}");
                    Dictionary<string, object> entry = BuildEntryFromRange(doc, code, range, isTarget: true);
                    if (entry != null)
                    {
                        result.Entries.Add(entry);
                    }
                    else
                    {
                        result.Errors.Add($"无法读取格式: {code}");
                    }
                }
            }

            result.Success = result.Entries.Count > 0;
            if (!result.Success)
            {
                result.Error = "未能组装任何 entries";
            }

            DbgLog(
                $"BuildContext targets=[{string.Join(",", targetCodes)}] " +
                $"entries={result.Entries.Count} warnings={result.Errors.Count}");
            return result;
        }

        public static List<string> GetOrderedParagraphCodes()
        {
            List<string> fromDisplay = ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay();
            if (fromDisplay.Count > 0)
            {
                return fromDisplay;
            }

            var codes = new List<string>();
            foreach (IReadOnlyList<string> row in DocumentState.ParagraphNameMapping)
            {
                if (row == null || row.Count == 0)
                {
                    continue;
                }

                string code = row[0]?.Trim();
                if (!string.IsNullOrEmpty(code) && code.StartsWith("P_", StringComparison.Ordinal))
                {
                    codes.Add(code);
                }
            }

            return codes;
        }

        private static Dictionary<string, object> BuildParagraphEntryFromRange(
            string paragraphCode,
            Word.Range paragraphRange,
            bool isTarget)
        {
            string content = DocumentState.GetParagraphContent(paragraphCode);
            if (string.IsNullOrEmpty(content) || paragraphRange == null)
            {
                return null;
            }

            IReadOnlyList<string> sentenceCodes = ParagraphCodeResolver.ExpandParagraphToSentenceCodes(paragraphCode);
            return new Dictionary<string, object>
            {
                ["paragraph_code"] = paragraphCode,
                ["text_preview"] = BuildTextPreview(content),
                ["is_target"] = isTarget,
                ["para_format"] = ParaFormatReader.Extract(paragraphRange),
                ["sentence_codes_in_paragraph"] = sentenceCodes.ToList(),
            };
        }

        private static Dictionary<string, object> BuildEntryFromRange(
            Word.Document doc,
            string code,
            Word.Range range,
            bool isTarget)
        {
            string content = DocumentState.GetSentenceContent(code);
            if (string.IsNullOrEmpty(content) || range == null)
            {
                return null;
            }

            Dictionary<string, object> snapshot = FormatInheritHelper.ExtractSnapshot(range);
            SemanticRoleHelper.RoleAnalysis role = SemanticRoleHelper.Analyze(code, content, snapshot);
            Word.Range paragraphRange = range.Paragraphs?.Count > 0 ? range.Paragraphs[1].Range : range;
            string paragraphCode = ParagraphCodeResolver.ResolveParagraphFromSentence(code) ?? "";
            var entry = new Dictionary<string, object>
            {
                ["code"] = code,
                ["paragraph_code"] = paragraphCode,
                ["text_preview"] = BuildTextPreview(content),
                ["is_target"] = isTarget,
                ["format_fingerprint"] = FormatInheritHelper.BuildFingerprint(snapshot),
                ["char_format"] = PickFields(snapshot, CharFormatKeys),
                ["para_format"] = ParaFormatReader.Extract(paragraphRange)
            };
            SemanticRoleHelper.AppendRoleFields(entry, role);
            DbgLog(
                $"entry {code} style_role={role.StyleRole} text_role={role.TextRole} " +
                $"semantic_role={role.SemanticRole ?? "(null)"} conflict={role.RoleConflict}");

            if (isTarget)
            {
                Dictionary<string, object> preReplace = DocumentState.TryGetPreReplaceFormat(code);
                if (preReplace != null && preReplace.Count > 0)
                {
                    entry["pre_replace_format"] = BuildPreReplaceFormat(preReplace);
                }
            }

            return entry;
        }

        public static List<string> GetOrderedSentenceCodes()
        {
            string snapshot = DocumentState.Snapshot ?? "";
            if (string.IsNullOrWhiteSpace(snapshot))
            {
                return new List<string>();
            }

            return snapshot
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s) && s.StartsWith("S_", StringComparison.Ordinal))
                .ToList();
        }

        private static Dictionary<string, object> BuildPreReplaceFormat(Dictionary<string, object> snapshot)
        {
            var pre = new Dictionary<string, object>
            {
                ["format_fingerprint"] = FormatInheritHelper.BuildFingerprint(snapshot),
                ["char_format"] = PickFields(snapshot, CharFormatKeys),
                ["para_format"] = PickFields(snapshot, ParaFormatKeys)
            };

            foreach (string key in CharFormatKeys.Concat(ParaFormatKeys))
            {
                if (snapshot.TryGetValue(key, out object value))
                {
                    pre[key] = value;
                }
            }

            return pre;
        }

        private static Dictionary<string, object> PickFields(
            Dictionary<string, object> snapshot,
            IEnumerable<string> keys)
        {
            var picked = new Dictionary<string, object>();
            if (snapshot == null)
            {
                return picked;
            }

            foreach (string key in keys)
            {
                if (snapshot.TryGetValue(key, out object value))
                {
                    picked[key] = value;
                }
            }

            return picked;
        }

        private static string BuildTextPreview(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return "";
            }

            string text = content.Replace("\r", " ").Replace("\n", " ").Trim();
            if (text.Length <= DefaultTextPreviewLength)
            {
                return text;
            }

            return text.Substring(0, DefaultTextPreviewLength);
        }

        public static List<string> ParseSentenceCodes(string codesCsv)
        {
            if (string.IsNullOrWhiteSpace(codesCsv))
            {
                return new List<string>();
            }

            return codesCsv
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }

        public static List<string> ParseParagraphCodes(string codesCsv)
        {
            if (string.IsNullOrWhiteSpace(codesCsv))
            {
                return new List<string>();
            }

            return codesCsv
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }

        public static bool IsParagraphCode(string code)
        {
            return !string.IsNullOrEmpty(code) && code.StartsWith("P_", StringComparison.Ordinal);
        }

        public static bool IsSentenceCode(string code)
        {
            return !string.IsNullOrEmpty(code) && code.StartsWith("S_", StringComparison.Ordinal);
        }
    }
}
