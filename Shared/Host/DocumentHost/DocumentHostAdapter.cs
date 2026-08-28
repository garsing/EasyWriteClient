using System;
using System.Collections.Generic;
using WordAddIn1.HostPlatform;
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
        /// <param name="activateDocument">
        /// 默认 false：不调用 Document.Activate，避免把 Word/WPS 抢到前台。
        /// </param>
        public static bool TryResolveContext(
            Dictionary<string, object> args,
            object wordApplication,
            out DocumentSessionContext context,
            out ToolResult errorResult,
            bool activateDocument = false)
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
                if (activateDocument)
                {
                    try
                    {
                        document.Activate();
                    }
                    catch (Exception)
                    {
                    }
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
        /// 解析渠道并得到可跑现网 Word Interop 管线的 Document（Word 原生；WPS 为兼容 RCW）。
        /// 供改字等仍大量依赖 <see cref="Word.Document"/> 的 F_* 渐进迁入；F_* 内勿写宿主分支。
        /// </summary>
        /// <param name="activateDocument">
        /// 默认 false：不激活文档窗口。写工具的前台由 Desktop 编排处理。
        /// </param>
        public static bool TryResolveInteropDocument(
            Dictionary<string, object> args,
            object wordApplication,
            out InteropDocumentHandle handle,
            out ToolResult errorResult,
            bool activateDocument = false)
        {
            handle = null;
            if (!TryResolveContext(
                    args,
                    wordApplication,
                    out DocumentSessionContext context,
                    out errorResult,
                    activateDocument))
            {
                return false;
            }

            if (context.Kind == ChannelKind.Word)
            {
                if (context.WordDocument == null)
                {
                    errorResult = new ToolResult
                    {
                        Success = false,
                        Error = "渠道对应的 Word 文档已关闭: " + context.ChannelId
                    };
                    return false;
                }

                handle = new InteropDocumentHandle(
                    context.WordDocument,
                    context,
                    disallowBackendApi: false);
                return true;
            }

            if (context.Kind == ChannelKind.Wps)
            {
                if (!DocumentHostCompat.TryPrepareWpsInteropDocument(
                        context,
                        out Word.Document wordDoc,
                        out string wpsError,
                        activateDocument))
                {
                    errorResult = new ToolResult
                    {
                        Success = false,
                        Error = string.IsNullOrEmpty(wpsError)
                            ? UnsupportedResult(ChannelKind.Wps, context.ChannelId, "interop_document").Error
                            : wpsError
                    };
                    return false;
                }

                handle = new InteropDocumentHandle(
                    wordDoc,
                    context,
                    disallowBackendApi: true);
                return true;
            }

            errorResult = UnsupportedResult(context.Kind, context.ChannelId, "interop_document");
            return false;
        }

        /// <summary>
        /// 试点：F_get_document_content 读文档 display 内容（只读，不 Activate 文档窗口）。
        /// </summary>
        public static bool TryGetDocumentContent(
            Dictionary<string, object> args,
            object wordApplication,
            string codeLevel,
            out GetDocumentContentHostResult result,
            out ToolResult errorResult)
        {
            result = null;
            // 只读：保持 Desktop 在前台，勿把 Word/WPS 抢到最前
            if (!TryResolveContext(
                    args,
                    wordApplication,
                    out DocumentSessionContext context,
                    out errorResult,
                    activateDocument: false))
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

        public static bool TryOpen(
            string fullPath,
            string app,
            bool createBlank,
            object wordApplication,
            out OpenDocumentResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (!OfficeOrWpsResolver.TryResolve(fullPath, app, out OfficeOrWps vendor, out string resolveError))
            {
                errorResult = new ToolResult { Success = false, Error = resolveError };
                return false;
            }

            bool ok;
            string error;
            if (vendor == OfficeOrWps.Office)
            {
                ok = WordDocumentHost.TryOpen(fullPath, createBlank, wordApplication, out result, out error);
            }
            else
            {
                ok = WpsDocumentHost.TryOpen(fullPath, createBlank, out result, out error);
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static ToolResult UnsupportedResult(ChannelKind kind, string channelId, string operation)
        {
            string host = kind.ToString().ToLowerInvariant();
            return new ToolResult
            {
                Success = false,
                Error =
                    (kind == ChannelKind.Excel || kind == ChannelKind.Et)
                        ? $"unsupported: 当前渠道是 {host}，读/改字工具尚未提供"
                          + (string.IsNullOrEmpty(channelId) ? "" : "（channel_id=" + channelId + "）")
                        : $"unsupported: 当前渠道宿主为 {host}，操作 {operation} 尚未接入文档适配层"
                          + (string.IsNullOrEmpty(channelId) ? "" : "（channel_id=" + channelId + "）")
            };
        }
    }
}
