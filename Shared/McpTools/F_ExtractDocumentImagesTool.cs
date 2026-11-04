using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;

namespace WordAddIn1
{
    /// <summary>
    /// 抽取当前文档图片：详细版 ProcessDocument，display 中带图片绝对路径。
    /// </summary>
    public static class F_ExtractDocumentImagesTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_extract_document_images"] = async (args) =>
            {
                try
                {
                    if (!F_GetDocumentContentTool.TryParseCodeLevel(args, out string codeLevel, out string codeLevelError))
                    {
                        return new ToolResult { Success = false, Error = codeLevelError };
                    }

                    bool isParagraph = codeLevel == F_GetDocumentContentTool.CodeLevelParagraph;

                    AgentRunCancellation.ThrowIfCancelled();

                    if (!DocumentHostAdapter.TryExtractDocumentImages(
                            args,
                            wordApplication,
                            codeLevel,
                            out GetDocumentContentHostResult hostResult,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    var processingResult = hostResult.ProcessingResult;
                    string documentContent = F_GetDocumentContentTool.BuildDocumentContent(processingResult, isParagraph);
                    string note = F_GetDocumentContentTool.BuildNote(processingResult.ProcessingMethod, isParagraph);
                    string duplicateCodeNote = isParagraph
                        ? F_GetDocumentContentTool.BuildDuplicateParagraphCodeNote()
                        : F_GetDocumentContentTool.BuildDuplicateSentenceCodeNote();

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
                        { "channel_id", ChannelRegistry.ToPublicId(hostResult.Context?.ChannelId) ?? "" },
                        { "host", hostResult.Context?.Kind.ToString().ToLowerInvariant() ?? "" },
                        { "message", "抽取成功" }
                    };

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
                    return new ToolResult { Success = false, Error = $"抽取文档图片失败: {ex.Message}" };
                }
            };
        }
    }
}
