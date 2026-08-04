using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 开发测试：P_ → Word 段落 Range 定位（批次 0b）。
    /// 对当前文档 ProcessDocument 后，按 paragraph display 顺序逐 P_ 走
    /// ResolveParagraphSpan + SpanParagraphLocator（与生产 apply 路径一致）。
    /// </summary>
    public static class F_TestParagraphCodeLocator
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_test_paragraph_code_locator"] = async (args) =>
            {
                try
                {
                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document doc, out ToolResult resolveError))
                    {
                        return resolveError;
                    }
                    bool runProcessDocument = GetBoolArg(args, "run_process_document", defaultValue: true);
                    bool selectOnSuccess = GetBoolArg(args, "select_on_success", defaultValue: false);
                    bool allowAutoTableScope = GetBoolArg(args, "allow_auto_table_scope", defaultValue: false);

                    LogHeader(doc.Name ?? "未命名文档");
                    Console.WriteLine(
                        allowAutoTableScope
                            ? "定位模式：自动限表（allow_auto_table_scope=true）"
                            : "定位模式：全文 Find（与句子测试一致，allow_auto_table_scope=false）");

                    WordDocumentExtractor.ProcessingResult processResult = null;
                    if (runProcessDocument)
                    {
                        processResult = WordDocumentExtractor.ProcessDocument(
                            doc,
                            new ProcessDocumentOptions { VerboseDebug = true });

                        if (processResult?.TableIndexToIdMap != null
                            && processResult.TableIndexToIdMap.Count > 0)
                        {
                            DocumentState.SetTableIndexToIdMap(processResult.TableIndexToIdMap);
                        }

                        Console.WriteLine(
                            $"ProcessDocument 完成：method={processResult?.ProcessingMethod}, " +
                            $"readText={processResult?.ReadText?.Length ?? 0}, " +
                            $"S_={processResult?.NameToContentMap?.Count ?? 0}, " +
                            $"P_={DocumentState.ParagraphNameMapping.Count}");
                    }
                    else if (DocumentState.ParagraphNameMapping.Count == 0)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "DocumentState 无 P_ mapping；请先 ProcessDocument 或设 run_process_document=true",
                        };
                    }

                    LocateTestSummary summary = RunParagraphLocateTests(doc, selectOnSuccess, allowAutoTableScope);
                    await Task.CompletedTask;

                    bool allOk = summary.LocateFailCount == 0 && summary.MappingMissingCount == 0;
                    return new ToolResult
                    {
                        Success = allOk,
                        Error = allOk
                            ? null
                            : $"P_ 定位失败 {summary.LocateFailCount} 处，mapping 缺失 {summary.MappingMissingCount} 处",
                        Data = summary.ToData(processResult, runProcessDocument),
                    };
                }
                catch (Exception ex)
                {
                    string errorMsg = $"测试 P_ 定位失败: {ex.Message}";
                    Console.WriteLine($"[ERROR] {errorMsg}");
                    Console.WriteLine(ex.StackTrace);
                    return new ToolResult { Success = false, Error = errorMsg };
                }
            };
        }

        private static LocateTestSummary RunParagraphLocateTests(
            Word.Document doc,
            bool selectOnSuccess,
            bool allowAutoTableScope)
        {
            var summary = new LocateTestSummary();
            List<string> orderedCodes = ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay();

            if (orderedCodes.Count == 0)
            {
                foreach (IReadOnlyList<string> row in DocumentState.ParagraphNameMapping)
                {
                    if (row != null && row.Count >= 1 && !string.IsNullOrEmpty(row[0]))
                    {
                        orderedCodes.Add(row[0]);
                    }
                }

                orderedCodes = orderedCodes.Distinct(StringComparer.Ordinal).ToList();
                Console.WriteLine(
                    $"paragraph display 为空，回退 unique P_ 列表（{orderedCodes.Count} 个，displayStart=0）");
            }
            else
            {
                Console.WriteLine($"paragraph display 有序 P_ 共 {orderedCodes.Count} 个（含重复出现）");
            }

            TableScopeIndex tableScope = TableParagraphScopeHelper.BuildIndexFromDocumentState();

            Console.WriteLine("");
            Console.WriteLine("=== P_ 定位测试（display 序 + ResolveParagraphSpan）===");
            int seq = 0;

            foreach (string code in orderedCodes)
            {
                seq++;
                int displayStart = seq - 1;
                string storedPreview = Truncate(
                    DocumentState.GetParagraphContent(code)?.Replace("\r", "\\r").Replace("\n", "\\n") ?? "",
                    80);

                string storedText = DocumentState.GetParagraphContent(code);
                string errorCode = null;
                string errorMessage = null;
                Word.Range range = null;

                if (string.IsNullOrEmpty(storedText))
                {
                    errorCode = ParagraphCodeResolver.ErrorParagraphMappingMissing;
                    errorMessage = $"未找到段落映射或 content 为空: {code}";
                }
                else if (ParagraphCodeAmbiguityHelper.IsEmptyParagraphStoredText(storedText))
                {
                    errorCode = ParagraphCodeResolver.ErrorParagraphRangeNotFound;
                    errorMessage = $"空段 {code} 无法定位 Word 段落 Range（I4-A）";
                }
                else if (doc == null)
                {
                    errorCode = ParagraphCodeResolver.ErrorParagraphMappingMissing;
                    errorMessage = "Word 文档实例不可用";
                }
                else
                {
                    string tableId = null;
                    if (allowAutoTableScope)
                    {
                        tableId = TableParagraphScopeHelper.ResolveEffectiveTableId(code, null, tableScope);
                    }

                    var codeList = new List<string> { code };
                    Word.Range span = DisplayPositionRangeResolver.ResolveParagraphSpan(
                        doc,
                        codeList,
                        displayStart,
                        tableId,
                        tableScope,
                        debugTag: code);

                    if (span == null)
                    {
                        errorCode = ParagraphCodeResolver.ErrorParagraphRangeNotFound;
                        errorMessage = $"无法定位段落 Span: {code} displayStart={displayStart}";
                    }
                    else
                    {
                        Word.Range hit = SpanParagraphLocator.LocateParagraph(
                            doc,
                            span,
                            code,
                            debugTag: code);
                        range = hit != null ? WordRangeFinder.ExpandToParagraphRange(hit) : null;

                        if (range == null)
                        {
                            errorCode = ParagraphCodeResolver.ErrorParagraphRangeNotFound;
                            errorMessage = $"Span 内未命中段落: {code} displayStart={displayStart}";
                        }
                    }
                }

                var item = new LocateTestItem
                {
                    Sequence = seq,
                    ParagraphCode = code,
                    DisplayStart = displayStart,
                    StoredPreview = storedPreview,
                    ErrorCode = errorCode,
                    ErrorMessage = errorMessage,
                };

                if (range != null)
                {
                    item.RangeStart = range.Start;
                    item.RangeEnd = range.End;
                    summary.LocateSuccessCount++;
                    Console.WriteLine(
                        $"  [{seq}] ✅ {code} displayStart={displayStart} Range={range.Start}-{range.End} text={storedPreview}");

                    if (selectOnSuccess)
                    {
                        try
                        {
                            range.Select();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"     （选中 Range 失败: {ex.Message}）");
                        }
                    }
                }
                else if (string.Equals(
                             errorCode,
                             ParagraphCodeResolver.ErrorParagraphRangeNotFound,
                             StringComparison.Ordinal)
                         && ParagraphCodeAmbiguityHelper.IsEmptyParagraphStoredText(
                             DocumentState.GetParagraphContent(code)))
                {
                    summary.EmptyParagraphCount++;
                    item.Outcome = "expected_empty";
                    Console.WriteLine(
                        $"  [{seq}] ⚪ {code} displayStart={displayStart} 空段（I4-A）→ {errorCode}: {errorMessage}");
                }
                else if (string.Equals(
                             errorCode,
                             ParagraphCodeResolver.ErrorParagraphMappingMissing,
                             StringComparison.Ordinal))
                {
                    summary.MappingMissingCount++;
                    item.Outcome = "mapping_missing";
                    summary.Failures.Add(item);
                    Console.WriteLine(
                        $"  [{seq}] ❌ {code} displayStart={displayStart} mapping 缺失: {errorMessage}");
                }
                else
                {
                    summary.LocateFailCount++;
                    item.Outcome = "locate_fail";
                    summary.Failures.Add(item);
                    Console.WriteLine(
                        $"  [{seq}] ❌ {code} displayStart={displayStart} → {errorCode}: {errorMessage} text={storedPreview}");
                }
            }

            Console.WriteLine("");
            Console.WriteLine(
                $"汇总：成功 {summary.LocateSuccessCount}，失败 {summary.LocateFailCount}，" +
                $"空段(I4-A) {summary.EmptyParagraphCount}，mapping 缺失 {summary.MappingMissingCount}");
            Console.WriteLine("");

            return summary;
        }

        private static bool GetBoolArg(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return defaultValue;
            }

            if (raw is bool b)
            {
                return b;
            }

            return bool.TryParse(raw.ToString(), out bool parsed) ? parsed : defaultValue;
        }

        private static void LogHeader(string documentName)
        {
            Console.WriteLine("");
            Console.WriteLine("==========================================");
            Console.WriteLine("=== F_test_paragraph_code_locator ===");
            Console.WriteLine("==========================================");
            Console.WriteLine($"文档: {documentName}");
        }

        private static string Truncate(string text, int maxLen)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLen)
            {
                return text ?? "";
            }

            return text.Substring(0, maxLen) + "...";
        }

        private sealed class LocateTestItem
        {
            public int Sequence { get; set; }
            public string ParagraphCode { get; set; }
            public int DisplayStart { get; set; }
            public string StoredPreview { get; set; }
            public string ErrorCode { get; set; }
            public string ErrorMessage { get; set; }
            public string Outcome { get; set; }
            public int RangeStart { get; set; }
            public int RangeEnd { get; set; }

            public Dictionary<string, object> ToDict()
            {
                var d = new Dictionary<string, object>
                {
                    ["sequence"] = Sequence,
                    ["paragraph_code"] = ParagraphCode,
                    ["display_start"] = DisplayStart,
                    ["outcome"] = Outcome ?? "success",
                    ["stored_preview"] = StoredPreview,
                };

                if (!string.IsNullOrEmpty(ErrorCode))
                {
                    d["error_code"] = ErrorCode;
                }

                if (!string.IsNullOrEmpty(ErrorMessage))
                {
                    d["error_message"] = ErrorMessage;
                }

                if (RangeStart > 0 || RangeEnd > 0)
                {
                    d["range_start"] = RangeStart;
                    d["range_end"] = RangeEnd;
                }

                return d;
            }
        }

        private sealed class LocateTestSummary
        {
            public int LocateSuccessCount { get; set; }
            public int LocateFailCount { get; set; }
            public int EmptyParagraphCount { get; set; }
            public int MappingMissingCount { get; set; }
            public List<LocateTestItem> Failures { get; } = new List<LocateTestItem>();

            public Dictionary<string, object> ToData(
                WordDocumentExtractor.ProcessingResult processResult,
                bool ranProcessDocument)
            {
                var data = new Dictionary<string, object>
                {
                    ["ran_process_document"] = ranProcessDocument,
                    ["display_paragraph_count"] =
                        ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay().Count,
                    ["unique_paragraph_mapping_count"] = DocumentState.ParagraphNameMapping.Count,
                    ["locate_success_count"] = LocateSuccessCount,
                    ["locate_fail_count"] = LocateFailCount,
                    ["empty_paragraph_count"] = EmptyParagraphCount,
                    ["mapping_missing_count"] = MappingMissingCount,
                    ["failures"] = Failures.Select(f => f.ToDict()).ToList(),
                };

                if (processResult != null)
                {
                    data["processing_method"] = processResult.ProcessingMethod;
                    data["read_text_length"] = processResult.ReadText?.Length ?? 0;
                    data["sentence_mappings_count"] = processResult.NameToContentMap?.Count ?? 0;
                }

                return data;
            }
        }
    }
}
