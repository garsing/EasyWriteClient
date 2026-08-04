using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
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

                    if (wordApplication == null)
                    {
                        return Task.FromResult(new ToolResult { Success = false, Error = "Word 应用程序不可用" });
                    }

                    if (!ChannelDocument.TryResolve(args, wordApplication, out _, out ToolResult resolveError))
                    {
                        return Task.FromResult(resolveError);
                    }

                    dynamic wordApp = wordApplication;
                    return ParagraphFormatApplyHelper.RunApplyAsync(wordApp, args);
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
    }
}
