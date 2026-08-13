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
            toolRegistry["F_apply_paragraph_format"] = (args) =>
            {
                try
                {
                    string logPath = EasyWriteLog.BeginSession("apply_paragraph_format");
                    FormatContextHelper.DbgLog($"会话日志: {logPath}");

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
                        return Task.FromResult(resolveError);
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_paragraph_format] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    Word.Application app = ResolveApplication(wordApplication, doc);
                    if (app == null)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "无法获取文档 Application（Word/WPS）"
                        });
                    }

                    return ParagraphFormatApplyHelper.RunApplyAsync(app, args, doc);
                }
                catch (Exception ex)
                {
                    FormatContextHelper.DbgLog($"F_apply_paragraph_format 异常: {ex.Message}");
                    return Task.FromResult(new ToolResult
                    {
                        Success = false,
                        Error = $"apply_paragraph_format 失败: {ex.Message}"
                    });
                }
            };
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
