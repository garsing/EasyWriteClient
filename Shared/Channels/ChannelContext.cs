using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// 从工具参数解析渠道：显式 channel_id → 否则默认渠道（D12）。
    /// B0：仅解析 id；B2 起返回 Word.Document。
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
    }
}
