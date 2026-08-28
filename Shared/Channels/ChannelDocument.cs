using System.Collections.Generic;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// F_* 工具入口：从参数/默认渠道解析 Word.Document（仅 Word 工具）。
    /// WPS 渠道返回明确 unsupported（尚未迁入适配层的工具）。
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

            if (TryGetResolvedNonWordFailure(args, out errorResult))
            {
                return false;
            }

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

            DocumentState.BindAndActivate(document);
            return true;
        }

        /// <summary>
        /// 显式/默认渠道为 WPS（或其它非 Word）时，未迁入工具应明确失败。
        /// </summary>
        public static bool TryGetResolvedNonWordFailure(
            Dictionary<string, object> args,
            out ToolResult errorResult)
        {
            errorResult = null;

            string explicitId = ChannelContext.TryGetChannelIdFromParameters(args);
            IOperationChannel channel = null;
            if (!string.IsNullOrEmpty(explicitId))
            {
                if (!ChannelRegistry.TryGet(explicitId, out channel))
                {
                    return false;
                }
            }
            else if (!ChannelRegistry.TryGetDefault(out channel) || channel == null)
            {
                return false;
            }

            if (channel.Kind == ChannelKind.Word)
            {
                return false;
            }

            errorResult = DocumentHostAdapter.UnsupportedResult(
                channel.Kind,
                channel.ChannelId,
                "word_only_tool");
            return true;
        }
    }
}
