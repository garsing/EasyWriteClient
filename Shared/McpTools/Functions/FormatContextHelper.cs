using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1.DocumentMapping;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    /// <summary>
    /// 格式上下文组装。字符格式来自 CharFormatRunReader，不附 pre_replace_format。
    /// </summary>
    public static class FormatContextHelper
    {
        public const string DbgPrefix = "[格式上下文][DBG]";
        public const int DefaultTextPreviewLength = 40;
        public const int DefaultContextWindow = 3;

        /// <summary>
        /// I19：先统一 Locator 出 Range，再判能否抽格式。null = 继续现网 S_ 抽格式。
        /// </summary>
        public static string LocateThenFormatPolicyError(
            Word.Document doc,
            string csv,
            bool disallowBackendApi)
        {
            if (string.IsNullOrWhiteSpace(csv))
            {
                return "缺少或无效的 target_codes";
            }

            var dict = new Dictionary<string, object> { ["codes"] = csv.Trim() };
            InsertTargetLocatorHelper.LocateSpanBatchResult batch = InsertTargetLocatorHelper.TryLocateSpans(
                doc,
                new[] { dict },
                disallowBackendApi);
            InsertTargetLocatorHelper.LocateSpanItem item = batch.Items.Count > 0 ? batch.Items[0] : null;
            if (batch.BatchBlocked || item == null || !item.Success)
            {
                return item?.Error ?? batch.Error ?? "locate_failed";
            }

            if (item.Spec != null && item.Spec.Kind == InsertTargetLocatorHelper.LocatorKind.EmptyCell)
            {
                return EmptyCellCode.NoFormatError;
            }

            return null;
        }

        public static void DbgLog(string message)
        {
            if (!FormatInheritHelper.EnableDbg)
            {
                return;
            }

            System.Diagnostics.Debug.WriteLine($"{DbgPrefix} {message}");
        }

        /// <summary>
        /// 套格式前重刷 S_/P_ 映射与 Snapshot。建副本会丢掉易失快照，
        /// 不刷新则 CheckBatch 找不到刚读过的码（改字工具已有同一步）。
        /// </summary>
        public static string RefreshMapping(
            Word.Document doc,
            string source,
            bool disallowBackendApi = false)
        {
            if (doc == null)
            {
                return "Word 文档不可用";
            }

            try
            {
                string pdError = InsertTargetLocatorHelper.TryProcessDocumentForLocate(doc, disallowBackendApi);
                return pdError == null ? null : "刷新文档映射失败: " + pdError;
            }
            catch (Exception ex)
            {
                return "刷新文档映射失败: " + ex.Message;
            }
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
            AmbiguityCheckResult precheckResult,
            bool disallowBackendApi = false,
            TextSpan textSpan = null)
        {
            return BuildSentenceContext(
                doc, targetCodesCsv, contextWindow, tableId, tableScope, precheckResult, disallowBackendApi, textSpan);
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
            AmbiguityCheckResult precheckResult,
            bool disallowBackendApi = false,
            TextSpan textSpan = null)
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

            InsertTargetLocatorHelper.LocateSpanBatchResult spanBatch = InsertTargetLocatorHelper.LocateFormatSpan(
                doc,
                targetCodes,
                paragraph: true,
                tableId,
                textSpan,
                disallowBackendApi,
                InsertTargetLocatorHelper.LocatePrecheckStream.ParagraphFormatContext,
                new ApplyFormatScopeArgs { TableId = tableId },
                "F_get_format_context");
            InsertTargetLocatorHelper.LocateSpanItem spanLocate = spanBatch.BatchBlocked
                ? new InsertTargetLocatorHelper.LocateSpanItem { Success = false, Error = spanBatch.Error }
                : spanBatch.Items[0];
            if (!spanLocate.Success && TextSpanArg.IsDirectError(spanLocate.Error))
            {
                result.Success = false;
                result.Error = spanLocate.Error;
                return result;
            }

            Word.Range span = spanLocate.Success ? spanLocate.Range : null;
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
            AmbiguityCheckResult precheckResult,
            bool disallowBackendApi = false,
            TextSpan textSpan = null)
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
                if (code.StartsWith("T_", StringComparison.Ordinal) && !EmptyCellCode.IsEmptyCellCode(code))
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

            InsertTargetLocatorHelper.LocateSpanBatchResult spanBatch = InsertTargetLocatorHelper.LocateFormatSpan(
                doc,
                targetCodes,
                paragraph: false,
                tableId,
                textSpan,
                disallowBackendApi,
                InsertTargetLocatorHelper.LocatePrecheckStream.SentenceFormatContext,
                new ApplyFormatScopeArgs { TableId = tableId },
                "F_get_format_context");
            InsertTargetLocatorHelper.LocateSpanItem spanLocate = spanBatch.BatchBlocked
                ? new InsertTargetLocatorHelper.LocateSpanItem { Success = false, Error = spanBatch.Error }
                : spanBatch.Items[0];
            if (!spanLocate.Success && TextSpanArg.IsDirectError(spanLocate.Error))
            {
                result.Success = false;
                result.Error = spanLocate.Error;
                return result;
            }

            Word.Range span = spanLocate.Success ? spanLocate.Range : null;
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
                    targetCodes[0],
                    span,
                    isTarget: true,
                    out string readError);
                if (singleEntry != null)
                {
                    result.Entries.Add(singleEntry);
                }
                else
                {
                    result.Errors.Add(readError ?? $"无法读取格式: {targetCodes[0]}");
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
                    Dictionary<string, object> entry = BuildEntryFromRange(code, range, isTarget: true, out string readError);
                    if (entry != null)
                    {
                        result.Entries.Add(entry);
                    }
                    else
                    {
                        result.Errors.Add(readError ?? $"无法读取格式: {code}");
                    }
                }
            }

            result.Success = result.Entries.Count > 0;
            if (!result.Success)
            {
                result.Error = result.Errors.Count > 0 ? result.Errors[0] : "未能组装任何 entries";
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
            var entry = new Dictionary<string, object>
            {
                ["paragraph_code"] = paragraphCode,
                ["text_preview"] = BuildTextPreview(content),
                ["is_target"] = isTarget,
                ["para_format"] = ParaFormatReader.Extract(paragraphRange),
                ["sentence_codes_in_paragraph"] = sentenceCodes.ToList(),
            };
            FormatLayerLookup.ApplyParagraph(paragraphRange.Document, paragraphCode, entry);
            return entry;
        }

        private static string ParagraphStyleName(Word.Range range)
        {
            try
            {
                Word.Style style = range?.get_Style() as Word.Style;
                return style?.NameLocal ?? "";
            }
            catch (Exception ex)
            {
                DbgLog("paragraph_style 读取失败: " + ex.Message);
                return "";
            }
        }

        private static Dictionary<string, object> BuildEntryFromRange(
            string code,
            Word.Range range,
            bool isTarget,
            out string error)
        {
            error = null;
            string content = DocumentState.GetSentenceContent(code);
            if (string.IsNullOrEmpty(content) || range == null)
            {
                error = "无法读取格式: " + code;
                return null;
            }

            if (!CharFormatRunReader.TryRead(range, out List<Dictionary<string, object>> runs, out error))
            {
                return null;
            }

            var roleSnap = new Dictionary<string, object>
            {
                ["paragraph_style"] = ParagraphStyleName(range)
            };
            SemanticRoleHelper.RoleAnalysis role = SemanticRoleHelper.Analyze(code, content, roleSnap);
            Word.Range paragraphRange = range.Paragraphs?.Count > 0 ? range.Paragraphs[1].Range : range;
            string paragraphCode = ParagraphCodeResolver.ResolveParagraphFromSentence(code) ?? "";
            var entry = new Dictionary<string, object>
            {
                ["code"] = code,
                ["paragraph_code"] = paragraphCode,
                ["text_preview"] = BuildTextPreview(content),
                ["is_target"] = isTarget,
                ["runs"] = runs,
                ["para_format"] = ParaFormatReader.Extract(paragraphRange)
            };
            FormatLayerLookup.ApplyParagraph(paragraphRange.Document, paragraphCode, entry);
            SemanticRoleHelper.AppendRoleFields(entry, role);
            DbgLog(
                $"entry {code} style_role={role.StyleRole} text_role={role.TextRole} " +
                $"semantic_role={role.SemanticRole ?? "(null)"} conflict={role.RoleConflict}");

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
            return ParagraphPublicCode.TryParse(code, out _);
        }

        public static bool IsSentenceCode(string code)
        {
            return !string.IsNullOrEmpty(code) && code.StartsWith("S_", StringComparison.Ordinal);
        }
    }
}
