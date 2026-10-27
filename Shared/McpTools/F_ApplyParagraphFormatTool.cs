using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 段落版式套用。文档触点经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyParagraphFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_paragraph_format"] = async (args) =>
            {
                try
                {
                    string logPath = EasyWriteLog.BeginSession("apply_paragraph_format");
                    FormatContextHelper.DbgLog($"会话日志: {logPath}");

                    if (ParagraphFormatApplyHelper.TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_apply_paragraph_format] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_paragraph_format] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    Word.Application app = ResolveApplication(wordApplication, doc);
                    if (app == null)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "无法获取文档 Application（Word/WPS）"
                        };
                    }

                    return await RunParagraphFormatBatch(app, args, doc, docHandle.DisallowBackendApi);
                }
                catch (Exception ex)
                {
                    FormatContextHelper.DbgLog($"F_apply_paragraph_format 异常: {ex.Message}");
                    return new ToolResult
                    {
                        Success = false,
                        Error = $"apply_paragraph_format 失败: {ex.Message}"
                    };
                }
            };
        }

        private static async Task<ToolResult> RunParagraphFormatBatch(
            Word.Application app,
            Dictionary<string, object> args,
            Word.Document doc,
            bool disallowBackendApi)
        {
            bool flatParagraph = args != null
                && args.ContainsKey("target_paragraph_codes")
                && !string.IsNullOrWhiteSpace(args["target_paragraph_codes"]?.ToString());
            if (!FormatApplyBatchCompiler.HasItems(args) && !flatParagraph)
            {
                return await ParagraphFormatApplyHelper.RunApplyAsync(app, args, doc, disallowBackendApi);
            }

            FormatApplyBatchCompiler.CompileResult compiled = FormatApplyBatchCompiler.CompileParagraph(args);
            if (!string.IsNullOrEmpty(compiled.ContractError))
            {
                return new ToolResult { Success = false, Error = compiled.ContractError };
            }

            if (!FormatApplyBatchCompiler.HasItems(args)
                && compiled.Ops.Count == 1
                && compiled.Warnings.Count == 0)
            {
                return await ParagraphFormatApplyHelper.RunApplyAsync(app, args, doc, disallowBackendApi);
            }

            var applied = new List<object>();
            var warnings = new List<object>();
            foreach (Dictionary<string, object> warning in compiled.Warnings)
            {
                warnings.Add(warning);
            }

            foreach (FormatApplyBatchCompiler.CompiledOp op in compiled.Ops)
            {
                Dictionary<string, object> oneArgs = FormatApplyBatchCompiler.CopyArgsForOp(args, op, "para_format");
                ToolResult one = await ParagraphFormatApplyHelper.RunApplyAsync(
                    app, oneArgs, doc, disallowBackendApi);
                string codes = string.Join(",", op.Codes ?? new List<string>());
                if (one != null && one.Success)
                {
                    applied.Add(new Dictionary<string, object>
                    {
                        ["item_index"] = op.ItemIndex,
                        ["codes"] = codes,
                        ["data"] = one.Data
                    });
                }
                else
                {
                    warnings.Add(FormatApplyBatchCompiler.Warning(
                        op.ItemIndex,
                        codes,
                        one?.Error ?? "无法定位"));
                }
            }

            var data = new Dictionary<string, object>
            {
                ["applied"] = applied,
                ["warnings"] = warnings
            };
            if (applied.Count == 0)
            {
                data["checkpoint_cleanup_only"] = true;
                string error = "未能套用任何格式";
                if (warnings.Count > 0 && warnings[0] is Dictionary<string, object> first && first.ContainsKey("message"))
                {
                    error = first["message"]?.ToString() ?? error;
                }

                return new ToolResult { Success = false, Error = error, Data = data };
            }

            return new ToolResult { Success = true, Data = data };
        }

        private static Word.Application ResolveApplication(object wordApplication, Word.Document doc)
        {
            // 渠道文档所属 Application 优先，避免 Word/WPS 双开时落到注入的 Word.ActiveDocument
            try
            {
                Word.Application fromDoc = doc?.Application;
                if (fromDoc != null)
                {
                    return fromDoc;
                }
            }
            catch (Exception)
            {
            }

            return wordApplication as Word.Application;
        }
    }
}
