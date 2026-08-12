using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    /// <summary>
    /// 读取文档 display 内容工具（支持 S_ 句子级与 P_ 段落级）。
    /// 文档触点经 <see cref="DocumentHostAdapter"/>，按 channel_id 分发 Word/WPS。
    /// </summary>
    public static class F_GetDocumentContentTool
    {
        public const string CodeLevelSentence = "sentence";
        public const string CodeLevelParagraph = "paragraph";

        /// <summary>
        /// 注册读取文档 display 内容工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_document_content"] = async (args) =>
            {
                try
                {
                    if (!TryParseCodeLevel(args, out string codeLevel, out string codeLevelError))
                    {
                        return new ToolResult { Success = false, Error = codeLevelError };
                    }

                    bool isParagraph = codeLevel == CodeLevelParagraph;

                    AgentRunCancellation.ThrowIfCancelled();

                    if (!DocumentHostAdapter.TryGetDocumentContent(
                            args,
                            wordApplication,
                            codeLevel,
                            out GetDocumentContentHostResult hostResult,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    var processingResult = hostResult.ProcessingResult;
                    string documentContent = BuildDocumentContent(processingResult, isParagraph);
                    string note = BuildNote(processingResult.ProcessingMethod, isParagraph);
                    string duplicateCodeNote = isParagraph
                        ? BuildDuplicateParagraphCodeNote()
                        : BuildDuplicateSentenceCodeNote();

                    if (!string.IsNullOrEmpty(duplicateCodeNote))
                    {
                        note = string.IsNullOrEmpty(note) ? duplicateCodeNote : note + "\n\n" + duplicateCodeNote;
                    }

                    var resultData = new Dictionary<string, object>
                    {
                        { "document_content", documentContent },
                        { "code_level", codeLevel },
                        { "processing_method", processingResult.ProcessingMethod ?? "Unknown" },
                        { "note", note },
                        { "channel_id", hostResult.Context?.ChannelId ?? "" },
                        { "host", hostResult.Context?.Kind.ToString().ToLowerInvariant() ?? "" }
                    };

                    System.Diagnostics.Debug.WriteLine(
                        $"[F_get_document_content] host={resultData["host"]}, channel_id={resultData["channel_id"]}, " +
                        $"code_level={codeLevel}, {resultData["processing_method"]}, " +
                        $"chunks={processingResult.ChunkInfo?.Count ?? 0}, " +
                        $"displayLen={documentContent.Length}");

                    if (EasyWriteDiagnostics.IsEnabled(DebugCategory.DocumentExtract))
                    {
                        System.Diagnostics.Debug.WriteLine("");
                        System.Diagnostics.Debug.WriteLine("==========================================");
                        System.Diagnostics.Debug.WriteLine("=== F_get_document_content 工具返回 ===");
                        System.Diagnostics.Debug.WriteLine("==========================================");
                        System.Diagnostics.Debug.WriteLine($"文档名称: {hostResult.DocumentDisplayName ?? "未命名文档"}");
                        System.Diagnostics.Debug.WriteLine($"channel_id: {resultData["channel_id"]}");
                        System.Diagnostics.Debug.WriteLine($"host: {resultData["host"]}");
                        System.Diagnostics.Debug.WriteLine($"code_level: {codeLevel}");
                        System.Diagnostics.Debug.WriteLine($"处理方式: {resultData["processing_method"]}");
                        System.Diagnostics.Debug.WriteLine($"返回内容长度: {documentContent.Length} 字符");
                        System.Diagnostics.Debug.WriteLine("==========================================");
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = resultData
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"读取文档内容失败: {ex.Message}" };
                }
            };
        }

        private static bool TryParseCodeLevel(
            Dictionary<string, object> args,
            out string codeLevel,
            out string error)
        {
            codeLevel = CodeLevelSentence;
            error = null;

            if (args == null || !args.ContainsKey("code_level") || args["code_level"] == null)
            {
                return true;
            }

            string raw = args["code_level"].ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(raw))
            {
                return true;
            }

            switch (raw.ToLowerInvariant())
            {
                case "sentence":
                case "s":
                case "s_":
                    codeLevel = CodeLevelSentence;
                    return true;
                case "paragraph":
                case "p":
                case "p_":
                    codeLevel = CodeLevelParagraph;
                    return true;
                default:
                    error = "code_level 无效，须为 sentence（S_，默认）或 paragraph（P_）";
                    return false;
            }
        }

        private static string BuildDocumentContent(WordDocumentExtractor.ProcessingResult processingResult, bool isParagraph)
        {
            if (processingResult.ProcessingMethod == "Local")
            {
                if (isParagraph)
                {
                    var paraDisplays = DocumentState.ChunkParagraphDisplayContents;
                    if (paraDisplays != null && paraDisplays.Count > 0)
                    {
                        return paraDisplays[0] ?? "";
                    }

                    return "";
                }

                if (processingResult.ChunkInfo != null && processingResult.ChunkInfo.Count > 0)
                {
                    return processingResult.ChunkInfo[0].DisplayContent ?? "";
                }

                return "";
            }

            if (processingResult.ProcessingMethod == "API")
            {
                var summaryParts = new List<string>();
                if (isParagraph)
                {
                    summaryParts.Add(
                        "⚠️ 重要提示：以下内容是各分块的摘要（Summary），不是段落级原文（paragraph_display_content）。"
                        + "如需获取完整 P_ display，请使用 F_get_curr_doc_chunk_paragraph_display_content 工具，输入具体的 chunk_idx。");
                }
                else
                {
                    summaryParts.Add(
                        "⚠️ 重要提示：以下内容是各分块的摘要（Summary），不是文档的原文内容（DisplayContent）。"
                        + "如需获取原文内容，请使用 F_get_curr_doc_chunk_display_content 工具。");
                }

                summaryParts.Add("");

                if (processingResult.ChunkInfo != null && processingResult.ChunkInfo.Count > 0)
                {
                    for (int i = 0; i < processingResult.ChunkInfo.Count; i++)
                    {
                        var chunkInfo = processingResult.ChunkInfo[i];
                        string summary = chunkInfo.Summary ?? "";
                        if (!string.IsNullOrEmpty(summary))
                        {
                            summaryParts.Add($"第{chunkInfo.ChunkIdx}个分块的摘要：{summary}");
                        }
                        else
                        {
                            summaryParts.Add($"第{chunkInfo.ChunkIdx}个分块的摘要：（无摘要）");
                        }
                    }
                }

                return string.Join("\n\n", summaryParts);
            }

            return "";
        }

        private static string BuildNote(string processingMethod, bool isParagraph)
        {
            if (processingMethod == "API")
            {
                if (isParagraph)
                {
                    return "⚠️ 当前返回的是各分块摘要，不是 paragraph_display_content。"
                        + "大文档请用 F_get_curr_doc_chunk_paragraph_display_content（chunk_idx）取完整 P_ 原文。"
                        + "\n\n[P_00000-start] 与 [P_00000-end] 是段落编号标记，不是正文。"
                        + "段版式精迁（F_apply_kb_paragraph_format）须使用 P_ 编码，与 S_ 不可混用。";
                }

                return "⚠️ 当前返回的是各分块摘要，不是 DisplayContent。"
                    + "大文档请用 F_get_curr_doc_chunk_display_content（chunk_idx）取完整 S_ 原文。"
                    + "\n\n[S_00000-start] 与 [S_00000-end] 是句子编号标记，不是正文。"
                    + "F_process_document_actions 的 new 中不要包含这些标记。";
            }

            if (isParagraph)
            {
                return "注意：[P_00000-start] 与 [P_00000-end] 是段落编号标记，不是正文。"
                    + "段版式精迁（F_apply_kb_paragraph_format）须使用 P_ 编码；句子改字仍用 S_（code_level=sentence）。"
                    + "若同一 P_ 出现多次，apply 时须读 suggested_locator_codes 写入 target_paragraph_codes（见 document-format §3.2）；禁止 index。";
            }

            return "注意：[S_00000-start] 与 [S_00000-end] 是句子编号标记，不是正文。"
                + "F_process_document_actions 定位时 new 中不要包含这些标记。"
                + " 若文档几乎为空（仅空 S_00000、无实质正文），insert 请省略 pre/post，在用户光标处插入，勿锚 S_00000。";
        }

        private static string BuildDuplicateSentenceCodeNote()
        {
            string snapshot = DocumentState.Snapshot;
            if (string.IsNullOrEmpty(snapshot))
            {
                return "";
            }

            var codes = DocumentState.GetOrderedSentenceCodesFromSnapshot(snapshot);
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (string code in codes)
            {
                if (!counts.ContainsKey(code))
                {
                    counts[code] = 0;
                }

                counts[code]++;
            }

            var duplicateCodes = counts
                .Where(kv => kv.Value >= 2)
                .Select(kv => kv.Key)
                .OrderBy(c => c, StringComparer.Ordinal)
                .ToList();

            if (duplicateCodes.Count == 0)
            {
                return "";
            }

            return "注意：以下句子编码在文档中出现多次（相同文本复用同一编码）："
                + string.Join(", ", duplicateCodes)
                + "。使用 F_process_document_actions 定位时，须在 original/deletion/pre/post 的 codes 中写入 suggested_locator_codes（见 document-format §3.3）；禁止 index。";
        }

        private static string BuildDuplicateParagraphCodeNote()
        {
            var codes = ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay();
            if (codes == null || codes.Count == 0)
            {
                return "";
            }

            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (string code in codes)
            {
                if (!counts.ContainsKey(code))
                {
                    counts[code] = 0;
                }

                counts[code]++;
            }

            var duplicateCodes = counts
                .Where(kv => kv.Value >= 2)
                .Select(kv => kv.Key)
                .OrderBy(c => c, StringComparer.Ordinal)
                .ToList();

            if (duplicateCodes.Count == 0)
            {
                return "";
            }

            return "注意：以下段落编码在文档中出现多次（相同段落文本复用同一编码）："
                + string.Join(", ", duplicateCodes)
                + "。使用 F_apply_paragraph_format / F_apply_kb_paragraph_format 定位时，须在 target_paragraph_codes 中写入 suggested_locator_codes（见 document-format §3.2）；禁止 index。";
        }
    }
}
