using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// 文档宿主适配层公用入口：按最终 channel_id 分发 Word / WPS（D1/D2/D11）。
    /// </summary>
    public static class DocumentHostAdapter
    {
        /// <summary>
        /// 解析渠道并构建会话上下文。不在此按宿主写 F_* 业务分支。
        /// </summary>
        public static bool TryResolveContext(
            Dictionary<string, object> args,
            object wordApplication,
            out DocumentSessionContext context,
            out ToolResult errorResult)
        {
            context = null;
            errorResult = null;

            if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string error))
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = string.IsNullOrEmpty(error) ? "无法解析操作渠道" : error
                };
                return false;
            }

            if (channel.Kind == ChannelKind.Word)
            {
                var wordChannel = channel as WordChannel;
                if (wordChannel == null || !wordChannel.TryGetLiveDocument(out Word.Document document))
                {
                    errorResult = new ToolResult
                    {
                        Success = false,
                        Error = "渠道对应的 Word 文档已关闭: " + channel.ChannelId
                    };
                    return false;
                }

                // Desktop 无注入时附着 Word（与 ChannelDocument 对齐）
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

                DocumentState.BindAndActivate(document);
                try
                {
                    document.Activate();
                }
                catch (Exception)
                {
                }

                context = new DocumentSessionContext(
                    wordChannel.ChannelId,
                    ChannelKind.Word,
                    wordChannel.DocUuid,
                    document,
                    null);
                return true;
            }

            if (channel.Kind == ChannelKind.Wps)
            {
                var wpsChannel = channel as WpsChannel;
                if (wpsChannel == null || !wpsChannel.TryGetLiveDocument(out object wpsDoc))
                {
                    errorResult = new ToolResult
                    {
                        Success = false,
                        Error = "渠道对应的 WPS 文档已关闭: " + channel.ChannelId
                    };
                    return false;
                }

                DocumentState.ActivateSessionByUuid(wpsChannel.DocUuid);
                context = new DocumentSessionContext(
                    wpsChannel.ChannelId,
                    ChannelKind.Wps,
                    wpsChannel.DocUuid,
                    null,
                    wpsDoc);
                return true;
            }

            errorResult = UnsupportedResult(channel.Kind, channel.ChannelId, "resolve_context");
            return false;
        }

        /// <summary>
        /// 试点：F_get_document_content 读文档 display 内容。
        /// </summary>
        public static bool TryGetDocumentContent(
            Dictionary<string, object> args,
            object wordApplication,
            string codeLevel,
            out GetDocumentContentHostResult result,
            out ToolResult errorResult)
        {
            result = null;
            if (!TryResolveContext(args, wordApplication, out DocumentSessionContext context, out errorResult))
            {
                return false;
            }

            try
            {
                if (context.Kind == ChannelKind.Word)
                {
                    result = WordDocumentHost.GetDocumentContent(context, codeLevel);
                    return true;
                }

                if (context.Kind == ChannelKind.Wps)
                {
                    result = WpsDocumentHost.GetDocumentContent(context, codeLevel);
                    return true;
                }

                errorResult = UnsupportedResult(context.Kind, context.ChannelId, "get_document_content");
                return false;
            }
            catch (InvalidOperationException ex)
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = ex.Message
                };
                return false;
            }
            catch (Exception ex)
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = $"读取文档内容失败: {ex.Message}"
                };
                return false;
            }
        }

        public static ToolResult UnsupportedResult(ChannelKind kind, string channelId, string operation)
        {
            string host = kind.ToString().ToLowerInvariant();
            return new ToolResult
            {
                Success = false,
                Error =
                    $"unsupported: 当前渠道宿主为 {host}，操作 {operation} 尚未接入文档适配层"
                    + (string.IsNullOrEmpty(channelId) ? "" : "（channel_id=" + channelId + "）")
            };
        }
    }
}
