using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1.OpenFiles;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1
{
    /// <summary>
    /// 进程级操作渠道注册表。channel_id 就是对外那一个号（w1/x1/p1/b1）。
    /// 已开稿按 (Kind, 路径/未保存键) 复用，不因 COM 包装更换而新发号。
    /// </summary>
    public static class ChannelRegistry
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, IOperationChannel> Channels =
            new Dictionary<string, IOperationChannel>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> KeyToChannelId =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> DocUuidToChannelId =
            new Dictionary<string, string>(StringComparer.Ordinal);
        private static readonly Dictionary<string, string> ChannelIdToDocUuid =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static int _nextW;
        private static int _nextX;
        private static int _nextP;
        private static int _nextB;
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

        /// <summary>默认渠道的 channel_id（w1/x1/p1/b1）。</summary>
        public static string PublicDefaultChannelId
        {
            get
            {
                lock (Gate)
                {
                    return string.IsNullOrEmpty(_defaultChannelId) ? null : _defaultChannelId;
                }
            }
        }

        /// <summary>恒等：ChannelId 已是对外号。</summary>
        public static string ToPublicId(string channelId)
        {
            return string.IsNullOrWhiteSpace(channelId) ? channelId : channelId.Trim();
        }

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
            string path = NormalizeDocumentPath(filePath) ?? NormalizeDocumentPath(WordChannel.TryReadFullName(doc));
            string unsavedName = TryReadComName(() => doc.Name) ?? "文档1";

            lock (Gate)
            {
                if (TryFindExistingUnlocked(
                        ChannelKind.Word,
                        path,
                        unsavedName,
                        LivePathOfWord,
                        out IOperationChannel existing)
                    && existing is WordChannel wordChannel)
                {
                    wordChannel.UpdateDocument(doc, path ?? filePath);
                    BindAfterUpdateUnlocked(wordChannel, path, unsavedName, uuid);
                    ClaimDefaultIfNeededUnlocked(wordChannel.ChannelId, claimDefaultIfEmpty);
                    return wordChannel;
                }

                string channelId = NextChannelIdUnlocked(ChannelKind.Word);
                var created = new WordChannel(channelId, uuid, doc, path ?? filePath);
                RegisterNewUnlocked(created, path, unsavedName, uuid, claimDefaultIfEmpty);
                return created;
            }
        }

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
            string path = NormalizeDocumentPath(filePath)
                ?? NormalizeDocumentPath(WpsCom.TryReadFullName(wpsDocument));
            string unsavedName = TryReadComName(() => WpsCom.TryReadName(wpsDocument)) ?? "文档1";

            lock (Gate)
            {
                if (TryFindExistingUnlocked(
                        ChannelKind.Wps,
                        path,
                        unsavedName,
                        LivePathOfWps,
                        out IOperationChannel existing)
                    && existing is WpsChannel wpsChannel)
                {
                    wpsChannel.UpdateDocument(wpsDocument, path ?? filePath);
                    BindAfterUpdateUnlocked(wpsChannel, path, unsavedName, uuid);
                    ClaimDefaultIfNeededUnlocked(wpsChannel.ChannelId, claimDefaultIfEmpty);
                    return wpsChannel;
                }

                string channelId = NextChannelIdUnlocked(ChannelKind.Wps);
                var created = new WpsChannel(channelId, uuid, wpsDocument, path ?? filePath);
                RegisterNewUnlocked(created, path, unsavedName, uuid, claimDefaultIfEmpty);
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
            string path = NormalizeDocumentPath(filePath)
                ?? NormalizeDocumentPath(ExcelChannel.TryReadFullName(workbook));
            string unsavedName = TryReadComName(() => workbook.Name) ?? "工作簿1";

            lock (Gate)
            {
                if (TryFindExistingUnlocked(
                        ChannelKind.Excel,
                        path,
                        unsavedName,
                        LivePathOfExcel,
                        out IOperationChannel existing)
                    && existing is ExcelChannel excelChannel)
                {
                    excelChannel.UpdateWorkbook(workbook, path ?? filePath);
                    BindAfterUpdateUnlocked(excelChannel, path, unsavedName, uuid);
                    ClaimDefaultIfNeededUnlocked(excelChannel.ChannelId, claimDefaultIfEmpty);
                    return excelChannel;
                }

                string channelId = NextChannelIdUnlocked(ChannelKind.Excel);
                var created = new ExcelChannel(channelId, uuid, workbook, path ?? filePath);
                RegisterNewUnlocked(created, path, unsavedName, uuid, claimDefaultIfEmpty);
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
            string path = NormalizeDocumentPath(filePath)
                ?? NormalizeDocumentPath(EtCom.TryReadFullName(workbook));
            string unsavedName = TryReadComName(() => EtCom.TryReadName(workbook)) ?? "工作簿1";

            lock (Gate)
            {
                if (TryFindExistingUnlocked(
                        ChannelKind.Et,
                        path,
                        unsavedName,
                        LivePathOfEt,
                        out IOperationChannel existing)
                    && existing is EtChannel etChannel)
                {
                    etChannel.UpdateWorkbook(workbook, path ?? filePath);
                    BindAfterUpdateUnlocked(etChannel, path, unsavedName, uuid);
                    ClaimDefaultIfNeededUnlocked(etChannel.ChannelId, claimDefaultIfEmpty);
                    return etChannel;
                }

                string channelId = NextChannelIdUnlocked(ChannelKind.Et);
                var created = new EtChannel(channelId, uuid, workbook, path ?? filePath);
                RegisterNewUnlocked(created, path, unsavedName, uuid, claimDefaultIfEmpty);
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
            string path = NormalizeDocumentPath(filePath)
                ?? NormalizeDocumentPath(PptChannel.TryReadFullName(presentation));
            string unsavedName = TryReadComName(() => presentation.Name) ?? "演示文稿1";

            lock (Gate)
            {
                if (TryFindExistingUnlocked(
                        ChannelKind.Ppt,
                        path,
                        unsavedName,
                        LivePathOfPpt,
                        out IOperationChannel existing)
                    && existing is PptChannel pptChannel)
                {
                    pptChannel.UpdatePresentation(presentation, path ?? filePath);
                    BindAfterUpdateUnlocked(pptChannel, path, unsavedName, uuid);
                    ClaimDefaultIfNeededUnlocked(pptChannel.ChannelId, claimDefaultIfEmpty);
                    return pptChannel;
                }

                string channelId = NextChannelIdUnlocked(ChannelKind.Ppt);
                var created = new PptChannel(channelId, uuid, presentation, path ?? filePath);
                RegisterNewUnlocked(created, path, unsavedName, uuid, claimDefaultIfEmpty);
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
            string path = NormalizeDocumentPath(filePath)
                ?? NormalizeDocumentPath(WppCom.TryReadFullName(presentation));
            string unsavedName = TryReadComName(() => WppCom.TryReadName(presentation)) ?? "演示文稿1";

            lock (Gate)
            {
                if (TryFindExistingUnlocked(
                        ChannelKind.Wpp,
                        path,
                        unsavedName,
                        LivePathOfWpp,
                        out IOperationChannel existing)
                    && existing is WppChannel wppChannel)
                {
                    wppChannel.UpdatePresentation(presentation, path ?? filePath);
                    BindAfterUpdateUnlocked(wppChannel, path, unsavedName, uuid);
                    ClaimDefaultIfNeededUnlocked(wppChannel.ChannelId, claimDefaultIfEmpty);
                    return wppChannel;
                }

                string channelId = NextChannelIdUnlocked(ChannelKind.Wpp);
                var created = new WppChannel(channelId, uuid, presentation, path ?? filePath);
                RegisterNewUnlocked(created, path, unsavedName, uuid, claimDefaultIfEmpty);
                return created;
            }
        }

        /// <summary>为易写浏览器页查找或创建渠道；channel_id 形如 <c>b1</c>。</summary>
        public static BrowserChannel CreateOrGetBrowserAgent(
            string tabUuid,
            bool setAsDefault = true)
        {
            return CreateOrGetBrowser(tabUuid, "agent", setAsDefault);
        }

        /// <summary>扩展附着页渠道；channel_id 形如 <c>b1</c>。</summary>
        public static BrowserChannel CreateOrGetBrowserAttach(
            string tabUuid,
            bool setAsDefault = true)
        {
            return CreateOrGetBrowser(tabUuid, "attach", setAsDefault);
        }

        private static BrowserChannel CreateOrGetBrowser(
            string tabUuid,
            string track,
            bool setAsDefault)
        {
            if (string.IsNullOrWhiteSpace(tabUuid))
            {
                throw new ArgumentException("tab_uuid 不能为空。", nameof(tabUuid));
            }

            string uuid = tabUuid.Trim();
            string trackNorm = string.IsNullOrWhiteSpace(track) ? "agent" : track.Trim();
            string docKey = BrowserDocumentKey(trackNorm, uuid);
            lock (Gate)
            {
                if (TryGetByDocumentKeyUnlocked(ChannelKind.Browser, docKey, out IOperationChannel existing)
                    && existing is BrowserChannel browserChannel
                    && string.Equals(browserChannel.Track, trackNorm, StringComparison.OrdinalIgnoreCase))
                {
                    if (setAsDefault)
                    {
                        _defaultChannelId = browserChannel.ChannelId;
                    }

                    RemapDocUuidUnlocked(browserChannel.ChannelId, uuid);
                    return browserChannel;
                }

                string channelId = NextChannelIdUnlocked(ChannelKind.Browser);
                var created = new BrowserChannel(channelId, uuid, trackNorm);
                Channels[channelId] = created;
                BindDocumentKeyUnlocked(channelId, ChannelKind.Browser, docKey);
                RemapDocUuidUnlocked(channelId, uuid);
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

        public static bool TryGetBrowserByTab(string tabUuid, string track, out BrowserChannel channel)
        {
            channel = null;
            if (string.IsNullOrWhiteSpace(tabUuid))
            {
                return false;
            }

            string trackNorm = string.IsNullOrWhiteSpace(track) ? "agent" : track.Trim();
            lock (Gate)
            {
                if (!TryGetByDocumentKeyUnlocked(
                        ChannelKind.Browser,
                        BrowserDocumentKey(trackNorm, tabUuid.Trim()),
                        out IOperationChannel ch)
                    || !(ch is BrowserChannel bc))
                {
                    return false;
                }

                channel = bc;
                return true;
            }
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
                string path = NormalizeDocumentPath(TryGetChannelFilePath(channel));
                string unsaved = string.IsNullOrEmpty(path)
                    ? (channel.TryGetDisplayNameSafe() ?? "未命名")
                    : null;
                if (channel is BrowserChannel browser)
                {
                    BindDocumentKeyUnlocked(
                        channel.ChannelId,
                        ChannelKind.Browser,
                        BrowserDocumentKey(browser.Track, browser.TabUuid));
                    RemapDocUuidUnlocked(channel.ChannelId, browser.TabUuid);
                }
                else
                {
                    RebindDocumentKeyUnlocked(channel.ChannelId, channel.Kind, path, unsaved);
                    string uuid = TryGetChannelDocUuid(channel);
                    if (!string.IsNullOrEmpty(uuid))
                    {
                        RemapDocUuidUnlocked(channel.ChannelId, uuid);
                    }
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
                return Channels.TryGetValue(channelId.Trim(), out channel);
            }
        }

        /// <summary>注册表里还有且 COM 仍活。侧栏漏扫时用这个决定能不能删渠。</summary>
        public static bool IsLive(string channelId)
        {
            if (!TryGet(channelId, out IOperationChannel ch) || ch == null)
            {
                return false;
            }

            if (ch is PptChannel ppt)
            {
                return ppt.TryGetLivePresentation(out _);
            }

            if (ch is WppChannel wpp)
            {
                return wpp.TryGetLivePresentation(out _);
            }

            if (ch is WordChannel word)
            {
                return word.TryGetLiveDocument(out _);
            }

            if (ch is WpsChannel wps)
            {
                return wps.TryGetLiveDocument(out _);
            }

            if (ch is ExcelChannel excel)
            {
                return excel.TryGetLiveWorkbook(out _);
            }

            if (ch is EtChannel et)
            {
                return et.TryGetLiveWorkbook(out _);
            }

            return ch is BrowserChannel;
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
                string key = channelId.Trim();
                if (!Channels.ContainsKey(key))
                {
                    return false;
                }

                _defaultChannelId = key;
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
                string key = channelId.Trim();
                if (!Channels.TryGetValue(key, out ch))
                {
                    return false;
                }

                channelId = ch.ChannelId;
                Channels.Remove(channelId);
                UnbindDocumentKeysUnlocked(channelId);
                UnmapDocUuidUnlocked(channelId);
                if (string.Equals(_defaultChannelId, channelId, StringComparison.OrdinalIgnoreCase))
                {
                    _defaultChannelId = null;
                }
            }

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

        public static IReadOnlyList<IOperationChannel> Snapshot()
        {
            lock (Gate)
            {
                return new List<IOperationChannel>(Channels.Values);
            }
        }

        public static void ClearAll()
        {
            lock (Gate)
            {
                Channels.Clear();
                KeyToChannelId.Clear();
                DocUuidToChannelId.Clear();
                ChannelIdToDocUuid.Clear();
                _nextW = 0;
                _nextX = 0;
                _nextP = 0;
                _nextB = 0;
                _defaultChannelId = null;
            }
        }

        internal static string NormalizeDocumentPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !IsSavedDocumentPath(path))
            {
                return null;
            }

            try
            {
                return Path.GetFullPath(path.Trim());
            }
            catch (Exception)
            {
                return path.Trim();
            }
        }

        private static bool IsSavedDocumentPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            try
            {
                return Path.IsPathRooted(path) && !string.IsNullOrEmpty(Path.GetDirectoryName(path));
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryFindExistingUnlocked(
            ChannelKind kind,
            string path,
            string unsavedName,
            Func<IOperationChannel, string> livePathOf,
            out IOperationChannel existing)
        {
            existing = null;
            if (!string.IsNullOrEmpty(path)
                && TryGetByDocumentKeyUnlocked(kind, path, out existing)
                && existing != null
                && existing.Kind == kind)
            {
                return true;
            }

            if (!string.IsNullOrEmpty(path))
            {
                foreach (IOperationChannel ch in Channels.Values)
                {
                    if (ch == null || ch.Kind != kind)
                    {
                        continue;
                    }

                    string stored = NormalizeDocumentPath(TryGetChannelFilePath(ch));
                    if (string.Equals(stored, path, StringComparison.OrdinalIgnoreCase))
                    {
                        existing = ch;
                        return true;
                    }

                    string live = livePathOf != null ? livePathOf(ch) : null;
                    if (string.Equals(live, path, StringComparison.OrdinalIgnoreCase))
                    {
                        existing = ch;
                        return true;
                    }
                }
            }

            if (string.IsNullOrEmpty(path) && !string.IsNullOrWhiteSpace(unsavedName)
                && TryGetByDocumentKeyUnlocked(kind, UnsavedDocumentKey(unsavedName), out existing)
                && existing != null
                && existing.Kind == kind)
            {
                return true;
            }

            existing = null;
            return false;
        }

        private static bool TryGetByDocumentKeyUnlocked(
            ChannelKind kind,
            string documentKey,
            out IOperationChannel channel)
        {
            channel = null;
            if (string.IsNullOrWhiteSpace(documentKey))
            {
                return false;
            }

            if (!KeyToChannelId.TryGetValue(ComposeKey(kind, documentKey), out string channelId))
            {
                return false;
            }

            return Channels.TryGetValue(channelId, out channel);
        }

        private static void RegisterNewUnlocked(
            IOperationChannel created,
            string path,
            string unsavedName,
            string docUuid,
            bool claimDefaultIfEmpty)
        {
            Channels[created.ChannelId] = created;
            RebindDocumentKeyUnlocked(created.ChannelId, created.Kind, path, unsavedName);
            RemapDocUuidUnlocked(created.ChannelId, docUuid);
            ClaimDefaultIfNeededUnlocked(created.ChannelId, claimDefaultIfEmpty);
        }

        private static void BindAfterUpdateUnlocked(
            IOperationChannel channel,
            string path,
            string unsavedName,
            string docUuid)
        {
            RebindDocumentKeyUnlocked(channel.ChannelId, channel.Kind, path, unsavedName);
            RemapDocUuidUnlocked(channel.ChannelId, docUuid);
        }

        private static void ClaimDefaultIfNeededUnlocked(string channelId, bool claimDefaultIfEmpty)
        {
            if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
            {
                _defaultChannelId = channelId;
            }
        }

        private static void RebindDocumentKeyUnlocked(
            string channelId,
            ChannelKind kind,
            string path,
            string unsavedName)
        {
            string key = !string.IsNullOrEmpty(path)
                ? path
                : UnsavedDocumentKey(string.IsNullOrWhiteSpace(unsavedName) ? "未命名" : unsavedName);
            BindDocumentKeyUnlocked(channelId, kind, key);
        }

        private static void BindDocumentKeyUnlocked(string channelId, ChannelKind kind, string documentKey)
        {
            UnbindDocumentKeysUnlocked(channelId);
            if (string.IsNullOrWhiteSpace(documentKey))
            {
                return;
            }

            KeyToChannelId[ComposeKey(kind, documentKey)] = channelId;
        }

        private static void UnbindDocumentKeysUnlocked(string channelId)
        {
            if (string.IsNullOrEmpty(channelId))
            {
                return;
            }

            var stale = new List<string>();
            foreach (KeyValuePair<string, string> kv in KeyToChannelId)
            {
                if (string.Equals(kv.Value, channelId, StringComparison.OrdinalIgnoreCase))
                {
                    stale.Add(kv.Key);
                }
            }

            for (int i = 0; i < stale.Count; i++)
            {
                KeyToChannelId.Remove(stale[i]);
            }
        }

        private static void RemapDocUuidUnlocked(string channelId, string docUuid)
        {
            UnmapDocUuidUnlocked(channelId);
            if (string.IsNullOrEmpty(channelId) || string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            DocUuidToChannelId[docUuid] = channelId;
            ChannelIdToDocUuid[channelId] = docUuid;
        }

        private static void UnmapDocUuidUnlocked(string channelId)
        {
            if (string.IsNullOrEmpty(channelId))
            {
                return;
            }

            if (ChannelIdToDocUuid.TryGetValue(channelId, out string old)
                && !string.IsNullOrEmpty(old))
            {
                DocUuidToChannelId.Remove(old);
            }

            ChannelIdToDocUuid.Remove(channelId);
        }

        private static string NextChannelIdUnlocked(ChannelKind kind)
        {
            char family = FamilyPrefix(kind);
            int n;
            switch (family)
            {
                case 'w':
                    n = ++_nextW;
                    break;
                case 'x':
                    n = ++_nextX;
                    break;
                case 'p':
                    n = ++_nextP;
                    break;
                default:
                    n = ++_nextB;
                    break;
            }

            return family + n.ToString();
        }

        private static char FamilyPrefix(ChannelKind kind)
        {
            switch (kind)
            {
                case ChannelKind.Word:
                case ChannelKind.Wps:
                    return 'w';
                case ChannelKind.Excel:
                case ChannelKind.Et:
                    return 'x';
                case ChannelKind.Ppt:
                case ChannelKind.Wpp:
                    return 'p';
                default:
                    return 'b';
            }
        }

        private static string ComposeKey(ChannelKind kind, string documentKey)
        {
            return kind.ToString() + "\n" + documentKey;
        }

        private static string UnsavedDocumentKey(string name)
        {
            return "unsaved:" + (name ?? "").Trim();
        }

        private static string BrowserDocumentKey(string track, string tabUuid)
        {
            return "browser:" + track + ":" + tabUuid;
        }

        private static string TryGetChannelFilePath(IOperationChannel channel)
        {
            if (channel is WordChannel word)
            {
                return word.FilePath;
            }

            if (channel is WpsChannel wps)
            {
                return wps.FilePath;
            }

            if (channel is ExcelChannel excel)
            {
                return excel.FilePath;
            }

            if (channel is EtChannel et)
            {
                return et.FilePath;
            }

            if (channel is PptChannel ppt)
            {
                return ppt.FilePath;
            }

            if (channel is WppChannel wpp)
            {
                return wpp.FilePath;
            }

            return null;
        }

        private static string TryGetChannelDocUuid(IOperationChannel channel)
        {
            if (channel is WordChannel word)
            {
                return word.DocUuid;
            }

            if (channel is WpsChannel wps)
            {
                return wps.DocUuid;
            }

            if (channel is ExcelChannel excel)
            {
                return excel.DocUuid;
            }

            if (channel is EtChannel et)
            {
                return et.DocUuid;
            }

            if (channel is PptChannel ppt)
            {
                return ppt.DocUuid;
            }

            if (channel is WppChannel wpp)
            {
                return wpp.DocUuid;
            }

            return null;
        }

        private static string TryReadComName(Func<string> read)
        {
            if (read == null)
            {
                return null;
            }

            try
            {
                string name = read();
                return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string LivePathOfWord(IOperationChannel channel)
        {
            return channel is WordChannel word && word.TryGetLiveDocument(out Word.Document doc)
                ? NormalizeDocumentPath(WordChannel.TryReadFullName(doc))
                : null;
        }

        private static string LivePathOfWps(IOperationChannel channel)
        {
            return channel is WpsChannel wps && wps.TryGetLiveDocument(out object doc)
                ? NormalizeDocumentPath(WpsCom.TryReadFullName(doc))
                : null;
        }

        private static string LivePathOfExcel(IOperationChannel channel)
        {
            return channel is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book)
                ? NormalizeDocumentPath(ExcelChannel.TryReadFullName(book))
                : null;
        }

        private static string LivePathOfEt(IOperationChannel channel)
        {
            return channel is EtChannel et && et.TryGetLiveWorkbook(out object book)
                ? NormalizeDocumentPath(EtCom.TryReadFullName(book))
                : null;
        }

        private static string LivePathOfPpt(IOperationChannel channel)
        {
            return channel is PptChannel ppt && ppt.TryGetLivePresentation(out PowerPoint.Presentation pres)
                ? NormalizeDocumentPath(PptChannel.TryReadFullName(pres))
                : null;
        }

        private static string LivePathOfWpp(IOperationChannel channel)
        {
            return channel is WppChannel wpp && wpp.TryGetLivePresentation(out object pres)
                ? NormalizeDocumentPath(WppCom.TryReadFullName(pres))
                : null;
        }
    }

    internal static class ChannelDisplayNameExtensions
    {
        public static string TryGetDisplayNameSafe(this IOperationChannel channel)
        {
            if (channel == null)
            {
                return null;
            }

            try
            {
                if (channel is WordChannel word)
                {
                    return word.TryGetDisplayName();
                }

                if (channel is WpsChannel wps)
                {
                    return wps.TryGetDisplayName();
                }

                if (channel is ExcelChannel excel)
                {
                    return excel.TryGetDisplayName();
                }

                if (channel is EtChannel et)
                {
                    return et.TryGetDisplayName();
                }

                if (channel is PptChannel ppt)
                {
                    return ppt.TryGetDisplayName();
                }

                if (channel is WppChannel wpp)
                {
                    return wpp.TryGetDisplayName();
                }

                if (channel is BrowserChannel browser)
                {
                    return browser.TryGetDisplayName();
                }
            }
            catch (Exception)
            {
            }

            return null;
        }
    }
}
