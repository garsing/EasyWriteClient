using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1
{
    /// <summary>
    /// 进程级操作渠道注册表。与 conversation 不强绑定（D15）。
    /// </summary>
    public static class ChannelRegistry
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, IOperationChannel> Channels =
            new Dictionary<string, IOperationChannel>(StringComparer.Ordinal);
        private static readonly Dictionary<string, string> DocUuidToChannelId =
            new Dictionary<string, string>(StringComparer.Ordinal);
        private static string _defaultChannelId;

        public static string DefaultChannelId
        {
            get
            {
                lock (Gate)
                {
                    return _defaultChannelId;
                }
            }
        }

        /// <summary>
        /// 为 Word 文档查找或创建渠道；channel_id 形如 <c>word:{doc_uuid}</c>。
        /// </summary>
        /// <param name="claimDefaultIfEmpty">
        /// 默认渠道为空时是否自动占用（默认 true，保持 Plugin / 旧工具行为）。
        /// 「打开文件」探测路径须传 false，避免抢默认渠道。
        /// </param>
        public static WordChannel CreateOrGetWord(
            Word.Document doc,
            string filePath = null,
            bool claimDefaultIfEmpty = true)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            string uuid = DocumentIdentity.EnsureUuid(doc);
            DocumentIdentity.EnsureCloseHandler(doc);

            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is WordChannel wordChannel)
                {
                    wordChannel.UpdateDocument(doc, filePath);
                    return wordChannel;
                }

                string channelId = "word:" + uuid;
                var created = new WordChannel(channelId, uuid, doc, filePath);
                Channels[channelId] = created;
                DocUuidToChannelId[uuid] = channelId;

                if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channelId;
                }

                return created;
            }
        }

        /// <summary>
        /// 为 WPS 文字文档查找或创建渠道；channel_id 形如 <c>wps:{doc_uuid}</c>。
        /// 禁止把 WPS 登记为 <c>word:</c> 渠道。
        /// </summary>
        /// <param name="claimDefaultIfEmpty">
        /// 默认渠道为空时是否自动占用。「打开文件」探测路径须传 false。
        /// </param>
        public static WpsChannel CreateOrGetWps(
            object wpsDocument,
            string filePath = null,
            bool claimDefaultIfEmpty = true)
        {
            if (wpsDocument == null)
            {
                throw new ArgumentNullException(nameof(wpsDocument));
            }

            string uuid = WpsDocumentIdentity.EnsureUuid(wpsDocument);

            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is WpsChannel wpsChannel)
                {
                    wpsChannel.UpdateDocument(wpsDocument, filePath);
                    return wpsChannel;
                }

                string channelId = "wps:" + uuid;
                var created = new WpsChannel(channelId, uuid, wpsDocument, filePath);
                Channels[channelId] = created;
                DocUuidToChannelId[uuid] = channelId;

                if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channelId;
                }

                return created;
            }
        }

        public static ExcelChannel CreateOrGetExcel(
            Excel.Workbook workbook,
            string filePath = null,
            bool claimDefaultIfEmpty = true)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            string uuid = ExcelWorkbookIdentity.EnsureUuid(workbook);
            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is ExcelChannel excelChannel)
                {
                    excelChannel.UpdateWorkbook(workbook, filePath);
                    return excelChannel;
                }

                string channelId = "excel:" + uuid;
                var created = new ExcelChannel(channelId, uuid, workbook, filePath);
                Channels[channelId] = created;
                DocUuidToChannelId[uuid] = channelId;
                if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channelId;
                }

                return created;
            }
        }

        public static EtChannel CreateOrGetEt(
            object workbook,
            string filePath = null,
            bool claimDefaultIfEmpty = true)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            string uuid = EtWorkbookIdentity.EnsureUuid(workbook);
            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is EtChannel etChannel)
                {
                    etChannel.UpdateWorkbook(workbook, filePath);
                    return etChannel;
                }

                string channelId = "et:" + uuid;
                var created = new EtChannel(channelId, uuid, workbook, filePath);
                Channels[channelId] = created;
                DocUuidToChannelId[uuid] = channelId;
                if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channelId;
                }

                return created;
            }
        }

        public static PptChannel CreateOrGetPpt(
            PowerPoint.Presentation presentation,
            string filePath = null,
            bool claimDefaultIfEmpty = true)
        {
            if (presentation == null)
            {
                throw new ArgumentNullException(nameof(presentation));
            }

            string uuid = PptPresentationIdentity.EnsureUuid(presentation);
            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is PptChannel pptChannel)
                {
                    pptChannel.UpdatePresentation(presentation, filePath);
                    return pptChannel;
                }

                string channelId = "ppt:" + uuid;
                var created = new PptChannel(channelId, uuid, presentation, filePath);
                Channels[channelId] = created;
                DocUuidToChannelId[uuid] = channelId;
                if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channelId;
                }

                return created;
            }
        }

        public static WppChannel CreateOrGetWpp(
            object presentation,
            string filePath = null,
            bool claimDefaultIfEmpty = true)
        {
            if (presentation == null)
            {
                throw new ArgumentNullException(nameof(presentation));
            }

            string uuid = WppPresentationIdentity.EnsureUuid(presentation);
            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is WppChannel wppChannel)
                {
                    wppChannel.UpdatePresentation(presentation, filePath);
                    return wppChannel;
                }

                string channelId = "wpp:" + uuid;
                var created = new WppChannel(channelId, uuid, presentation, filePath);
                Channels[channelId] = created;
                DocUuidToChannelId[uuid] = channelId;
                if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channelId;
                }

                return created;
            }
        }

        /// <summary>
        /// 为易写浏览器页查找或创建渠道；channel_id 形如 <c>browser:agent:{tab_uuid}</c>。
        /// </summary>
        public static BrowserChannel CreateOrGetBrowserAgent(
            string tabUuid,
            bool setAsDefault = true)
        {
            if (string.IsNullOrWhiteSpace(tabUuid))
            {
                throw new ArgumentException("tab_uuid 不能为空。", nameof(tabUuid));
            }

            string uuid = tabUuid.Trim();
            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is BrowserChannel browserChannel
                    && string.Equals(browserChannel.Track, "agent", StringComparison.OrdinalIgnoreCase))
                {
                    if (setAsDefault)
                    {
                        _defaultChannelId = browserChannel.ChannelId;
                    }

                    return browserChannel;
                }

                string channelId = BrowserChannel.AgentPrefix + uuid;
                var created = new BrowserChannel(channelId, uuid, "agent");
                Channels[channelId] = created;
                DocUuidToChannelId[uuid] = channelId;

                if (setAsDefault || string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channelId;
                }

                return created;
            }
        }

        public static bool TryGetBrowser(string channelId, out BrowserChannel channel)
        {
            channel = null;
            if (!TryGet(channelId, out IOperationChannel ch) || !(ch is BrowserChannel bc))
            {
                return false;
            }

            channel = bc;
            return true;
        }

        public static void Register(IOperationChannel channel, bool setAsDefault = false)
        {
            if (channel == null)
            {
                throw new ArgumentNullException(nameof(channel));
            }

            if (string.IsNullOrWhiteSpace(channel.ChannelId))
            {
                throw new ArgumentException("channel_id 不能为空。", nameof(channel));
            }

            lock (Gate)
            {
                Channels[channel.ChannelId] = channel;
                if (channel is WordChannel wc && !string.IsNullOrEmpty(wc.DocUuid))
                {
                    DocUuidToChannelId[wc.DocUuid] = channel.ChannelId;
                }
                else if (channel is WpsChannel wps && !string.IsNullOrEmpty(wps.DocUuid))
                {
                    DocUuidToChannelId[wps.DocUuid] = channel.ChannelId;
                }
                else if (channel is ExcelChannel excel && !string.IsNullOrEmpty(excel.DocUuid))
                {
                    DocUuidToChannelId[excel.DocUuid] = channel.ChannelId;
                }
                else if (channel is EtChannel et && !string.IsNullOrEmpty(et.DocUuid))
                {
                    DocUuidToChannelId[et.DocUuid] = channel.ChannelId;
                }
                else if (channel is PptChannel ppt && !string.IsNullOrEmpty(ppt.DocUuid))
                {
                    DocUuidToChannelId[ppt.DocUuid] = channel.ChannelId;
                }
                else if (channel is WppChannel wpp && !string.IsNullOrEmpty(wpp.DocUuid))
                {
                    DocUuidToChannelId[wpp.DocUuid] = channel.ChannelId;
                }
                else if (channel is BrowserChannel browser && !string.IsNullOrEmpty(browser.TabUuid))
                {
                    DocUuidToChannelId[browser.TabUuid] = channel.ChannelId;
                }

                if (setAsDefault || string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channel.ChannelId;
                }
            }
        }

        public static bool TryGet(string channelId, out IOperationChannel channel)
        {
            channel = null;
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            lock (Gate)
            {
                return Channels.TryGetValue(channelId, out channel);
            }
        }

        public static bool TryGetWord(string channelId, out WordChannel channel)
        {
            channel = null;
            if (!TryGet(channelId, out IOperationChannel ch) || !(ch is WordChannel wc))
            {
                return false;
            }

            channel = wc;
            return true;
        }

        public static bool TryGetWps(string channelId, out WpsChannel channel)
        {
            channel = null;
            if (!TryGet(channelId, out IOperationChannel ch) || !(ch is WpsChannel wc))
            {
                return false;
            }

            channel = wc;
            return true;
        }

        public static bool TryGetByDocUuid(string docUuid, out WordChannel channel)
        {
            channel = null;
            if (string.IsNullOrWhiteSpace(docUuid))
            {
                return false;
            }

            lock (Gate)
            {
                if (!DocUuidToChannelId.TryGetValue(docUuid, out string channelId))
                {
                    return false;
                }

                if (!Channels.TryGetValue(channelId, out IOperationChannel ch) || !(ch is WordChannel wc))
                {
                    return false;
                }

                channel = wc;
                return true;
            }
        }

        public static bool TryGetDefaultWps(out WpsChannel channel)
        {
            channel = null;
            if (!TryGetDefault(out IOperationChannel ch) || !(ch is WpsChannel wc))
            {
                return false;
            }

            channel = wc;
            return true;
        }

        public static bool SetDefault(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            lock (Gate)
            {
                if (!Channels.ContainsKey(channelId))
                {
                    return false;
                }

                _defaultChannelId = channelId;
                return true;
            }
        }

        public static IOperationChannel GetDefaultOrNull()
        {
            return TryGetDefault(out IOperationChannel channel) ? channel : null;
        }

        public static bool TryGetDefault(out IOperationChannel channel)
        {
            channel = null;
            lock (Gate)
            {
                if (string.IsNullOrEmpty(_defaultChannelId))
                {
                    return false;
                }

                return Channels.TryGetValue(_defaultChannelId, out channel);
            }
        }

        public static bool TryGetDefaultWord(out WordChannel channel)
        {
            channel = null;
            if (!TryGetDefault(out IOperationChannel ch) || !(ch is WordChannel wc))
            {
                return false;
            }

            channel = wc;
            return true;
        }

        /// <summary>
        /// Plugin：活动文档变化时查找/创建 Word 渠道并设为默认（§4.3）。
        /// </summary>
        public static WordChannel SyncDefaultFromActiveDocument(Word.Document document)
        {
            if (document == null)
            {
                return null;
            }

            WordChannel channel = CreateOrGetWord(document);
            SetDefault(channel.ChannelId);
            return channel;
        }

        public static bool Remove(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            IOperationChannel ch;
            lock (Gate)
            {
                if (!Channels.TryGetValue(channelId, out ch))
                {
                    return false;
                }

                Channels.Remove(channelId);
                if (ch is WordChannel wc && !string.IsNullOrEmpty(wc.DocUuid))
                {
                    DocUuidToChannelId.Remove(wc.DocUuid);
                }
                else if (ch is WpsChannel wps && !string.IsNullOrEmpty(wps.DocUuid))
                {
                    DocUuidToChannelId.Remove(wps.DocUuid);
                }
                else if (ch is ExcelChannel excel && !string.IsNullOrEmpty(excel.DocUuid))
                {
                    DocUuidToChannelId.Remove(excel.DocUuid);
                }
                else if (ch is EtChannel et && !string.IsNullOrEmpty(et.DocUuid))
                {
                    DocUuidToChannelId.Remove(et.DocUuid);
                }
                else if (ch is PptChannel ppt && !string.IsNullOrEmpty(ppt.DocUuid))
                {
                    DocUuidToChannelId.Remove(ppt.DocUuid);
                }
                else if (ch is WppChannel wpp && !string.IsNullOrEmpty(wpp.DocUuid))
                {
                    DocUuidToChannelId.Remove(wpp.DocUuid);
                }
                else if (ch is BrowserChannel browser && !string.IsNullOrEmpty(browser.TabUuid))
                {
                    DocUuidToChannelId.Remove(browser.TabUuid);
                }

                if (string.Equals(_defaultChannelId, channelId, StringComparison.Ordinal))
                {
                    _defaultChannelId = null;
                }
            }

            // 锁外释放 COM，避免 FinalRelease 回调再进 Registry
            if (ch is EtChannel etRelease)
            {
                etRelease.ReleaseCom();
            }
            else if (ch is WppChannel wppRelease)
            {
                wppRelease.ReleaseCom();
            }

            return true;
        }

        public static bool RemoveByDocUuid(string docUuid)
        {
            if (string.IsNullOrWhiteSpace(docUuid))
            {
                return false;
            }

            string channelId;
            lock (Gate)
            {
                if (!DocUuidToChannelId.TryGetValue(docUuid, out channelId))
                {
                    return false;
                }
            }

            return Remove(channelId);
        }

        /// <summary>测试/进程退出用。</summary>
        public static void ClearAll()
        {
            lock (Gate)
            {
                Channels.Clear();
                DocUuidToChannelId.Clear();
                _defaultChannelId = null;
            }
        }
    }
}
