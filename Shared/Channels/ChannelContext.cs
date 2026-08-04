using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从工具参数解析渠道：显式 channel_id → 否则默认渠道 →（Plugin）ActiveDocument（D12 / I6）。
    /// </summary>
    public static class ChannelContext
    {
        public const string ChannelIdParameterName = "channel_id";

        public static string TryGetChannelIdFromParameters(IDictionary<string, object> parameters)
        {
            if (parameters == null)
            {
                return null;
            }

            if (!parameters.TryGetValue(ChannelIdParameterName, out object raw) || raw == null)
            {
                return null;
            }

            string id = Convert.ToString(raw);
            return string.IsNullOrWhiteSpace(id) ? null : id.Trim();
        }

        /// <summary>
        /// 解析本次调用应使用的渠道。显式 channel_id 不改变默认渠道。
        /// </summary>
        public static bool TryResolveChannel(
            IDictionary<string, object> parameters,
            out IOperationChannel channel,
            out string error)
        {
            channel = null;
            error = null;

            string explicitId = TryGetChannelIdFromParameters(parameters);
            if (!string.IsNullOrEmpty(explicitId))
            {
                if (!ChannelRegistry.TryGet(explicitId, out channel))
                {
                    error = "未知 channel_id: " + explicitId;
                    return false;
                }

                return true;
            }

            if (!ChannelRegistry.TryGetDefault(out channel) || channel == null)
            {
                error = "无默认操作渠道；请先打开或新建文档（F_open_word_document）。";
                return false;
            }

            return true;
        }

        /// <summary>
        /// 解析 Word.Document：显式 channel_id → 默认渠道 → Plugin 下回退 ActiveDocument；
        /// 成功时 <see cref="DocumentState.BindAndActivate"/>。
        /// </summary>
        public static bool TryResolveWordDocument(
            IDictionary<string, object> parameters,
            object wordApplication,
            out Word.Document document,
            out string error)
        {
            document = null;
            error = null;

            string explicitId = TryGetChannelIdFromParameters(parameters);
            if (!string.IsNullOrEmpty(explicitId))
            {
                if (!ChannelRegistry.TryGetWord(explicitId, out WordChannel explicitChannel))
                {
                    error = "未知 channel_id: " + explicitId;
                    return false;
                }

                if (!explicitChannel.TryGetLiveDocument(out document))
                {
                    error = "渠道对应的 Word 文档已关闭: " + explicitId;
                    return false;
                }

                DocumentState.BindAndActivate(document);
                return true;
            }

            if (ChannelRegistry.TryGetDefaultWord(out WordChannel defaultChannel)
                && defaultChannel.TryGetLiveDocument(out document))
            {
                DocumentState.BindAndActivate(document);
                return true;
            }

            if (ChannelHost.AllowActiveDocumentFallback)
            {
                Word.Document active = TryGetActiveDocument(wordApplication);
                if (active != null)
                {
                    WordChannel created = ChannelRegistry.SyncDefaultFromActiveDocument(active);
                    if (created != null && created.TryGetLiveDocument(out document))
                    {
                        DocumentState.BindAndActivate(document);
                        return true;
                    }
                }
            }

            error = ChannelHost.Kind == ChannelHostKind.Desktop
                ? "无默认操作渠道；请先打开或新建文档（F_open_word_document）。"
                : "无可用 Word 文档；请先打开文档或指定 channel_id。";
            return false;
        }

        /// <summary>
        /// 同 <see cref="TryResolveWordDocument"/>，失败时抛 <see cref="InvalidOperationException"/>。
        /// </summary>
        public static Word.Document ResolveWordDocument(
            IDictionary<string, object> parameters,
            object wordApplication)
        {
            if (!TryResolveWordDocument(parameters, wordApplication, out Word.Document document, out string error))
            {
                throw new InvalidOperationException(error ?? "无法解析 Word 文档渠道。");
            }

            return document;
        }

        /// <summary>
        /// Plugin 活动窗口切换：同步默认渠道并激活 DocumentState 分片。
        /// </summary>
        public static void NotifyPluginActiveDocumentChanged(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            ChannelRegistry.SyncDefaultFromActiveDocument(document);
            DocumentState.OnActiveDocumentChanged(document);
        }

        private static Word.Document TryGetActiveDocument(object wordApplication)
        {
            if (wordApplication == null)
            {
                return null;
            }

            try
            {
                if (wordApplication is Word.Application app)
                {
                    return app.ActiveDocument;
                }

                // 动态/COM 晚绑定兜底
                dynamic dyn = wordApplication;
                return dyn.ActiveDocument as Word.Document;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
