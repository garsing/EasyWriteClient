using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// F_* 工具入口：从参数/默认渠道解析 Word.Document（B5）。
    /// </summary>
    public static class ChannelDocument
    {
        /// <summary>
        /// 成功时已 <see cref="DocumentState.BindAndActivate"/>。
        /// </summary>
        public static bool TryResolve(
            Dictionary<string, object> args,
            object wordApplication,
            out Word.Document document,
            out ToolResult errorResult)
        {
            document = null;
            errorResult = null;

            // Desktop 启动不注入 Word；F_open 之后同轮工具需附着已运行实例
            if (wordApplication == null)
            {
                if (!WordApplicationResolver.TryResolve(
                        null,
                        out Word.Application attached,
                        out string attachError,
                        createIfMissing: false))
                {
                    errorResult = new ToolResult
                    {
                        Success = false,
                        Error = string.IsNullOrEmpty(attachError)
                            ? "Word应用程序实例不可用"
                            : attachError
                    };
                    return false;
                }

                wordApplication = attached;
            }

            if (!ChannelContext.TryResolveWordDocument(
                    args,
                    wordApplication,
                    out document,
                    out string error))
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = string.IsNullOrEmpty(error) ? "无法解析 Word 文档渠道" : error
                };
                return false;
            }

            // 使 Word.Application.ActiveDocument 与渠道文档一致（供仍读 ActiveDocument 的 Helper）
            try
            {
                document.Activate();
            }
            catch (System.Exception)
            {
            }

            return true;
        }
    }
}
