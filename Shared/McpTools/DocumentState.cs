using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 句子级别的文档状态管理（按 doc_uuid 分片，活动窗口读写当前分片）。
    /// 使用S_前缀 + 5位编码名称（S_00000, S_00001, ..., S_00009, S_0000a, ..., S_0000z, S_0000A, ..., S_ZZZZZ）来跟踪句子
    /// 编码顺序：数字(0-9) -> 小写字母(a-z) -> 大写字母(A-Z)
    /// 使用T_前缀 + 哈希值前8位（如T_eStyRRUe）来跟踪表格，如果冲突则使用更长的哈希值
    /// </summary>
    public static class DocumentState
    {
        private static readonly Dictionary<string, DocumentSessionState> SessionsByUuid =
            new Dictionary<string, DocumentSessionState>(StringComparer.Ordinal);

        private static DocumentSessionState _activeSession;
        private static string _activeDocUuid = "";

        /// <summary>递增后清空 ProcessDocument 会话缓存（指纹算法或缓存条目结构变更时）。</summary>
        public const int FingerprintVersion = 6;

        /// <summary>当前活动文档分片；无绑定时为临时空分片（不写回 SessionsByUuid）。</summary>
        private static DocumentSessionState Session
        {
            get
            {
                if (_activeSession != null)
                {
                    return _activeSession;
                }

                return _scratchSession;
            }
        }

        private static readonly DocumentSessionState _scratchSession = new DocumentSessionState();

        internal static DocumentSessionState GetOrCreateSession(string docUuid)
        {
            if (string.IsNullOrEmpty(docUuid))
            {
                throw new ArgumentException("docUuid 不能为空", nameof(docUuid));
            }

            if (!SessionsByUuid.TryGetValue(docUuid, out DocumentSessionState session))
            {
                session = new DocumentSessionState { DocUuid = docUuid };
                SessionsByUuid[docUuid] = session;
            }

            return session;
        }

        /// <summary>
        /// 工具或 ProcessDocument 入口：绑定 Word 文档并切换为活动分片。
        /// </summary>
        public static void BindAndActivate(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            string uuid = DocumentIdentity.EnsureUuid(document);
            DocumentIdentity.EnsureCloseHandler(document);
            ActivateSession(uuid);
        }

        /// <summary>
        /// 用户切换活动窗口时载入该文档已绑定的 uuid 分片（未处理过的文档无分片）。
        /// </summary>
        public static void OnActiveDocumentChanged(Word.Document document)
        {
            if (document == null)
            {
                _activeSession = null;
                _activeDocUuid = "";
                return;
            }

            string uuid = DocumentIdentity.TryResolveUuid(document);
            if (string.IsNullOrEmpty(uuid))
            {
                _activeSession = null;
                _activeDocUuid = "";
                System.Diagnostics.Debug.WriteLine("[DocumentState] active document has no uuid binding yet");
                return;
            }

            ActivateSession(uuid);
        }

        private static void ActivateSession(string uuid)
        {
            if (string.IsNullOrEmpty(uuid))
            {
                _activeSession = null;
                _activeDocUuid = "";
                return;
            }

            _activeDocUuid = uuid;
            _activeSession = GetOrCreateSession(uuid);
            _activeSession.DocUuid = uuid;
        }

        /// <summary>
        /// 关文档：仅移除该文档 uuid 分片（D16）。
        /// </summary>
        public static void ClearOnDocumentClose(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            string uuid = DocumentIdentity.TryResolveUuid(document);
            DocumentIdentity.Unregister(document);

            if (!string.IsNullOrEmpty(uuid))
            {
                SessionsByUuid.Remove(uuid);
                ChannelRegistry.RemoveByDocUuid(uuid);
                if (string.Equals(_activeDocUuid, uuid, StringComparison.Ordinal))
                {
                    _activeSession = null;
                    _activeDocUuid = "";
                }

                System.Diagnostics.Debug.WriteLine($"[DocumentState] session removed on document close uuid={uuid}");
            }
        }

        /// <summary>
        /// 插件退出：清除全部分片与绑定。
        /// </summary>
        public static void ClearAllOnShutdown()
        {
            SessionsByUuid.Clear();
            _activeSession = null;
            _activeDocUuid = "";
            DocumentIdentity.ClearAll();
            DocumentCheckpointService.ClearAllOnShutdown();
            ChannelRegistry.ClearAll();
            System.Diagnostics.Debug.WriteLine("[DocumentState] all sessions cleared on shutdown");
        }

        /// <summary>
        /// checkpoint restore / mutating 失败回滚后：仅清易失状态（I6/B）；保留句子映射与表格映射。
        /// </summary>
        public static void ClearAfterCheckpointRestore()
        {
            if (_activeSession == null)
            {
                return;
            }

            _activeSession.ClearVolatileState();
            System.Diagnostics.Debug.WriteLine(
                $"[DocumentState] cleared volatile state after checkpoint restore uuid={_activeDocUuid} (mapping retained, I6/B)");
        }

        /// <summary>
        /// 用户「新建会话」：清除活动文档的 DocumentState 分片与 checkpoint；保留 DocumentIdentity 绑定。
        /// </summary>
        public static void ClearSessionForNewConversation(Word.Document document)
        {
            if (document == null)
            {
                _activeSession = null;
                _activeDocUuid = "";
                System.Diagnostics.Debug.WriteLine("[DocumentState] ClearSessionForNewConversation: no active document");
                return;
            }

            string uuid = DocumentIdentity.TryResolveUuid(document);
            if (!string.IsNullOrEmpty(uuid))
            {
                SessionsByUuid.Remove(uuid);
                DocumentCheckpointService.ClearSessionForDocument(uuid);
                System.Diagnostics.Debug.WriteLine(
                    $"[DocumentState] ClearSessionForNewConversation uuid={uuid} (identity retained)");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine(
                    "[DocumentState] ClearSessionForNewConversation: no uuid bound yet, skip Remove");
            }

            _activeSession = null;
            _activeDocUuid = "";
        }

        private static string GenerateContentHash(string content)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(content ?? ""));
                return Convert.ToBase64String(hashBytes);
            }
        }

        private static string GenerateNextName()
        {
            string name = "";
            int index = Session.NextNameIndex++;

            for (int i = 0; i < 5; i++)
            {
                int charIndex = index % 62;
                char ch;

                if (charIndex < 10)
                {
                    ch = (char)('0' + charIndex);
                }
                else if (charIndex < 36)
                {
                    ch = (char)('a' + (charIndex - 10));
                }
                else
                {
                    ch = (char)('A' + (charIndex - 36));
                }

                name = ch + name;
                index /= 62;
            }

            return "S_" + name;
        }

        private static string GenerateNextParagraphName()
        {
            string name = "";
            int index = Session.NextParagraphNameIndex++;

            for (int i = 0; i < 5; i++)
            {
                int charIndex = index % 62;
                char ch;

                if (charIndex < 10)
                {
                    ch = (char)('0' + charIndex);
                }
                else if (charIndex < 36)
                {
                    ch = (char)('a' + (charIndex - 10));
                }
                else
                {
                    ch = (char)('A' + (charIndex - 36));
                }

                name = ch + name;
                index /= 62;
            }

            return "P_" + name;
        }

        public static string FindOrCreateParagraphName(string storedText, int mark = 0)
        {
            string stored = storedText ?? "";
            string contentHash = GenerateContentHash(stored);

            foreach (var mapping in Session.ParagraphNameMapping)
            {
                if (mapping.Count >= 2 && mapping[1] == contentHash)
                {
                    if (mapping.Count >= 4)
                    {
                        mapping[3] = mark.ToString();
                    }
                    else
                    {
                        mapping.Add(mark.ToString());
                    }

                    return mapping[0];
                }
            }

            string newName = GenerateNextParagraphName();
            Session.ParagraphNameMapping.Add(new List<string> { newName, contentHash, stored, mark.ToString() });
            return newName;
        }

        public static string GetParagraphContent(string paragraphName)
        {
            foreach (var mapping in Session.ParagraphNameMapping)
            {
                if (mapping.Count >= 3 && mapping[0] == paragraphName)
                {
                    return mapping[2];
                }
            }

            return "";
        }

        public static void ClearSentenceParagraphIndex()
        {
            Session.SentenceToParagraph.Clear();
            Session.ParagraphToSentences.Clear();
        }

        public static void LinkSentenceToParagraph(string sentenceName, string paragraphName)
        {
            if (string.IsNullOrEmpty(sentenceName) || string.IsNullOrEmpty(paragraphName))
            {
                return;
            }

            Session.SentenceToParagraph[sentenceName] = paragraphName;
            if (!Session.ParagraphToSentences.TryGetValue(paragraphName, out List<string> bucket))
            {
                bucket = new List<string>();
                Session.ParagraphToSentences[paragraphName] = bucket;
            }

            if (!bucket.Contains(sentenceName))
            {
                bucket.Add(sentenceName);
            }
        }

        public static void SetChunkParagraphDisplayContents(IReadOnlyList<string> contents)
        {
            Session.ChunkParagraphDisplayContents.Clear();
            if (contents == null)
            {
                return;
            }

            foreach (string content in contents)
            {
                Session.ChunkParagraphDisplayContents.Add(content ?? "");
            }
        }

        public static void SetParagraphMappings(List<List<string>> paragraphMappings)
        {
            Session.ParagraphNameMapping.Clear();
            if (paragraphMappings != null)
            {
                foreach (var mapping in paragraphMappings)
                {
                    if (mapping != null && mapping.Count >= 4)
                    {
                        Session.ParagraphNameMapping.Add(new List<string>
                        {
                            mapping[0], mapping[1], mapping[2], mapping[3]
                        });
                    }
                    else if (mapping != null && mapping.Count >= 3)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[DocumentState] ⚠️ 警告：段落映射格式不正确，期望四元组，实际收到 {mapping.Count} 个元素，已跳过");
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[DocumentState] 已设置段落映射，共 {Session.ParagraphNameMapping.Count} 条");
        }

        public static void ApplySentenceParagraphIndexFromApi(
            Dictionary<string, string> sentenceToParagraph,
            Dictionary<string, List<string>> paragraphToSentences)
        {
            ClearSentenceParagraphIndex();
            if (sentenceToParagraph != null)
            {
                foreach (var kv in sentenceToParagraph)
                {
                    LinkSentenceToParagraph(kv.Key, kv.Value);
                }
            }

            if (paragraphToSentences != null)
            {
                foreach (var kv in paragraphToSentences)
                {
                    if (!Session.ParagraphToSentences.ContainsKey(kv.Key))
                    {
                        Session.ParagraphToSentences[kv.Key] = new List<string>();
                    }

                    foreach (string s in kv.Value ?? new List<string>())
                    {
                        if (!Session.ParagraphToSentences[kv.Key].Contains(s))
                        {
                            Session.ParagraphToSentences[kv.Key].Add(s);
                        }
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[DocumentState] 已应用 S_↔P_ 索引：{Session.SentenceToParagraph.Count} 条");
        }

        public static IReadOnlyList<IReadOnlyList<string>> ParagraphNameMapping => Session.ParagraphNameMapping;

        public static IReadOnlyList<IReadOnlyList<string>> SentenceNameMapping => Session.SentenceNameMapping;

        public static IReadOnlyDictionary<string, string> SentenceToParagraph => Session.SentenceToParagraph;

        public static IReadOnlyDictionary<string, List<string>> ParagraphToSentences => Session.ParagraphToSentences;

        public static IReadOnlyList<string> ChunkParagraphDisplayContents => Session.ChunkParagraphDisplayContents;

        public static string FindOrCreateSentenceName(string content, int mark = 0)
        {
            string storedContent = content ?? "";
            string contentHash = GenerateContentHash(
                TextUtils.PrepareForSentenceSplit(storedContent, SentenceSplitNewlineMode.Readback));

            foreach (var mapping in Session.SentenceNameMapping)
            {
                if (mapping.Count >= 2 && mapping[1] == contentHash)
                {
                    if (mapping.Count >= 4)
                    {
                        mapping[3] = mark.ToString();
                    }
                    else
                    {
                        mapping.Add(mark.ToString());
                    }

                    return mapping[0];
                }
            }

            string newName = GenerateNextName();
            Session.SentenceNameMapping.Add(new List<string> { newName, contentHash, storedContent, mark.ToString() });
            return newName;
        }

        public static string GetSentenceContent(string sentenceName)
        {
            foreach (var mapping in Session.SentenceNameMapping)
            {
                if (mapping.Count >= 3 && mapping[0] == sentenceName)
                {
                    return mapping[2];
                }
            }

            return "";
        }

        public static int GetSentenceMark(string sentenceName)
        {
            foreach (var mapping in Session.SentenceNameMapping)
            {
                if (mapping.Count >= 3 && mapping[0] == sentenceName)
                {
                    if (mapping.Count >= 4)
                    {
                        return int.TryParse(mapping[3], out int mark) ? mark : 0;
                    }

                    return 0;
                }
            }

            return 0;
        }

        public static string FindOrCreateTableId(string tableContent)
        {
            string normalizedContent = tableContent ?? "";
            string contentHash = GenerateContentHash(normalizedContent);

            foreach (var mapping in Session.TableIdMapping)
            {
                if (mapping.Count >= 2 && mapping[1] == contentHash)
                {
                    string storedId = mapping[0];
                    if (!storedId.StartsWith("T_"))
                    {
                        mapping[0] = "T_" + storedId;
                        return mapping[0];
                    }

                    return storedId;
                }
            }

            string tableIdBase = contentHash.Length >= 8 ? contentHash.Substring(0, 8) : contentHash;
            string tableIdWithPrefix = "T_" + tableIdBase;

            bool tableIdExists = false;
            foreach (var mapping in Session.TableIdMapping)
            {
                if (mapping.Count >= 1)
                {
                    string storedId = mapping[0];
                    if (storedId == tableIdWithPrefix || storedId == tableIdBase)
                    {
                        tableIdExists = true;
                        break;
                    }
                }
            }

            if (tableIdExists)
            {
                for (int len = 12; len <= contentHash.Length; len += 4)
                {
                    string candidateIdBase = contentHash.Substring(0, len);
                    string candidateIdWithPrefix = "T_" + candidateIdBase;
                    bool candidateExists = false;
                    foreach (var mapping in Session.TableIdMapping)
                    {
                        if (mapping.Count >= 1)
                        {
                            string storedId = mapping[0];
                            if (storedId == candidateIdWithPrefix || storedId == candidateIdBase)
                            {
                                candidateExists = true;
                                break;
                            }
                        }
                    }

                    if (!candidateExists)
                    {
                        tableIdBase = candidateIdBase;
                        tableIdWithPrefix = candidateIdWithPrefix;
                        break;
                    }
                }
            }

            Session.TableIdMapping.Add(new List<string> { tableIdWithPrefix, contentHash, normalizedContent });
            return tableIdWithPrefix;
        }

        public static string GetTableContent(string tableId)
        {
            if (string.IsNullOrEmpty(tableId))
            {
                return "";
            }

            foreach (var mapping in Session.TableIdMapping)
            {
                if (mapping.Count >= 3 && mapping[0] == tableId)
                {
                    return mapping[2];
                }
            }

            return "";
        }

        public static void AddTableIdToOrder(string tableId)
        {
            if (!string.IsNullOrEmpty(tableId))
            {
                Session.TableIdOrder.Add(tableId);
            }
        }

        public static int GetTableIndex(string tableId)
        {
            if (string.IsNullOrEmpty(tableId))
            {
                return -1;
            }

            return Session.TableIdOrder.IndexOf(tableId);
        }

        public static List<string> GetTableIdOrder()
        {
            return new List<string>(Session.TableIdOrder);
        }

        public static void ClearTableIdOrder()
        {
            Session.TableIdOrder.Clear();
        }

        public static void SetTableIndexToIdMap(Dictionary<int, string> tableIndexToIdMap)
        {
            Session.TableIdOrder.Clear();
            if (tableIndexToIdMap != null)
            {
                var orderedKeys = tableIndexToIdMap.Keys.OrderBy(k => k);
                foreach (var index in orderedKeys)
                {
                    if (!string.IsNullOrEmpty(tableIndexToIdMap[index]))
                    {
                        Session.TableIdOrder.Add(tableIndexToIdMap[index]);
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置表格序号到编号映射，共 {Session.TableIdOrder.Count} 个表格");
        }

        public static string FindOrCreateChartId(string chartBodyContent)
        {
            string normalizedContent = chartBodyContent ?? "";
            string contentHash = GenerateContentHash(normalizedContent);

            foreach (var mapping in Session.ChartIdMapping)
            {
                if (mapping.Count >= 2 && mapping[1] == contentHash)
                {
                    string storedId = mapping[0];
                    if (!storedId.StartsWith("C_"))
                    {
                        mapping[0] = "C_" + storedId;
                        return mapping[0];
                    }

                    return storedId;
                }
            }

            string chartIdBase = contentHash.Length >= 8 ? contentHash.Substring(0, 8) : contentHash;
            string chartIdWithPrefix = "C_" + chartIdBase;

            bool chartIdExists = false;
            foreach (var mapping in Session.ChartIdMapping)
            {
                if (mapping.Count >= 1)
                {
                    string storedId = mapping[0];
                    if (storedId == chartIdWithPrefix || storedId == chartIdBase)
                    {
                        chartIdExists = true;
                        break;
                    }
                }
            }

            if (chartIdExists)
            {
                for (int len = 12; len <= contentHash.Length; len += 4)
                {
                    string candidateIdBase = contentHash.Substring(0, len);
                    string candidateIdWithPrefix = "C_" + candidateIdBase;
                    bool candidateExists = false;
                    foreach (var mapping in Session.ChartIdMapping)
                    {
                        if (mapping.Count >= 1)
                        {
                            string storedId = mapping[0];
                            if (storedId == candidateIdWithPrefix || storedId == candidateIdBase)
                            {
                                candidateExists = true;
                                break;
                            }
                        }
                    }

                    if (!candidateExists)
                    {
                        chartIdBase = candidateIdBase;
                        chartIdWithPrefix = candidateIdWithPrefix;
                        break;
                    }
                }
            }

            Session.ChartIdMapping.Add(new List<string> { chartIdWithPrefix, contentHash, normalizedContent });
            return chartIdWithPrefix;
        }

        public static void AddChartIdToOrder(string chartId)
        {
            if (!string.IsNullOrEmpty(chartId))
            {
                Session.ChartIdOrder.Add(chartId);
            }
        }

        public static int GetChartIndex(string chartId)
        {
            if (string.IsNullOrEmpty(chartId))
            {
                return -1;
            }

            return Session.ChartIdOrder.IndexOf(chartId);
        }

        public static List<string> GetChartIdOrder()
        {
            return new List<string>(Session.ChartIdOrder);
        }

        public static void ClearChartIdOrder()
        {
            Session.ChartIdOrder.Clear();
        }

        public static void SetChartIdOrder(IReadOnlyList<string> chartIdOrder)
        {
            Session.ChartIdOrder.Clear();
            if (chartIdOrder != null)
            {
                foreach (string chartId in chartIdOrder)
                {
                    AddChartIdToOrder(chartId);
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置图表序号到编号映射，共 {Session.ChartIdOrder.Count} 个图表");
        }

        public static void SetChartIdMappingSnapshot(List<List<string>> mappingSnapshot)
        {
            Session.ChartIdMapping.Clear();
            if (mappingSnapshot == null)
            {
                return;
            }

            foreach (var mapping in mappingSnapshot)
            {
                if (mapping == null || mapping.Count == 0)
                {
                    continue;
                }

                Session.ChartIdMapping.Add(new List<string>(mapping));
            }
        }

        public static List<List<string>> GetChartIdMappingSnapshot()
        {
            var snapshot = new List<List<string>>();
            foreach (var mapping in Session.ChartIdMapping)
            {
                snapshot.Add(new List<string>(mapping));
            }

            return snapshot;
        }

        /// <summary>
        /// 按 drawing 元数据为图片实例分配唯一 I_（文档内禁止重复；同 meta 加盐消歧）。
        /// 与 Table/Chart「同 hash 复用」相反：每次调用都是新实例入口。
        /// </summary>
        public static string AllocateUniqueImageId(string meta)
        {
            string normalizedMeta = meta ?? "";
            var occupied = new HashSet<string>(StringComparer.Ordinal);
            foreach (var mapping in Session.ImageIdMapping)
            {
                if (mapping != null && mapping.Count >= 1 && !string.IsNullOrEmpty(mapping[0]))
                {
                    occupied.Add(mapping[0]);
                }
            }

            int k = 1;
            while (true)
            {
                // D15a：初算撞占用则 k=2,3,… 对 meta|#k 重哈希（同 meta 第二张须 k>=2）
                string material = (k == 1) ? normalizedMeta : $"{normalizedMeta}|#{k}";
                string fullHash = GenerateContentHash(material);
                int prefixLen = 8;
                string imageId = "I_" + (fullHash.Length >= prefixLen
                    ? fullHash.Substring(0, prefixLen)
                    : fullHash);

                if (!occupied.Contains(imageId))
                {
                    Session.ImageIdMapping.Add(new List<string>
                    {
                        imageId,
                        fullHash,
                        normalizedMeta,
                        k.ToString()
                    });
                    System.Diagnostics.Debug.WriteLine(
                        $"[ImageId] allocated {imageId} k={k} meta={TruncateForLog(normalizedMeta)}");
                    return imageId;
                }

                k++;
                if (k > 10000)
                {
                    throw new InvalidOperationException("AllocateUniqueImageId: 消歧盐耗尽");
                }
            }
        }

        private static string TruncateForLog(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            return value.Length <= 80 ? value : value.Substring(0, 80) + "...";
        }

        public static void AddImageIdToOrder(string imageId)
        {
            if (!string.IsNullOrEmpty(imageId))
            {
                Session.ImageIdOrder.Add(imageId);
            }
        }

        public static int GetImageIndex(string imageId)
        {
            if (string.IsNullOrEmpty(imageId))
            {
                return -1;
            }

            return Session.ImageIdOrder.IndexOf(imageId);
        }

        public static List<string> GetImageIdOrder()
        {
            return new List<string>(Session.ImageIdOrder);
        }

        public static void ClearImageIdOrder()
        {
            Session.ImageIdOrder.Clear();
        }

        public static void ClearImageIdMapping()
        {
            Session.ImageIdMapping.Clear();
        }

        public static void SetImageIdOrder(IReadOnlyList<string> imageIdOrder)
        {
            Session.ImageIdOrder.Clear();
            if (imageIdOrder != null)
            {
                foreach (string imageId in imageIdOrder)
                {
                    AddImageIdToOrder(imageId);
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[DocumentState] 已设置图片序号到编号映射，共 {Session.ImageIdOrder.Count} 个图片");
        }

        public static void SetImageIndexToIdMap(Dictionary<int, string> imageIndexToIdMap)
        {
            Session.ImageIdOrder.Clear();
            if (imageIndexToIdMap != null)
            {
                foreach (var index in imageIndexToIdMap.Keys.OrderBy(k => k))
                {
                    if (!string.IsNullOrEmpty(imageIndexToIdMap[index]))
                    {
                        Session.ImageIdOrder.Add(imageIndexToIdMap[index]);
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[DocumentState] 已设置图片序号到编号映射，共 {Session.ImageIdOrder.Count} 个图片");
        }

        public static void SetImageIdMappingSnapshot(List<List<string>> mappingSnapshot)
        {
            Session.ImageIdMapping.Clear();
            if (mappingSnapshot == null)
            {
                return;
            }

            foreach (var mapping in mappingSnapshot)
            {
                if (mapping == null || mapping.Count == 0)
                {
                    continue;
                }

                Session.ImageIdMapping.Add(new List<string>(mapping));
            }
        }

        public static List<List<string>> GetImageIdMappingSnapshot()
        {
            var snapshot = new List<List<string>>();
            foreach (var mapping in Session.ImageIdMapping)
            {
                snapshot.Add(new List<string>(mapping));
            }

            return snapshot;
        }

        public static string Snapshot => Session.Snapshot;

        public static string SegSnapshot => Session.SegSnapshot;

        public static IReadOnlyList<SnapshotRecord> SnapshotHistory => Session.SnapshotHistory;

        public static int LatestSnapshotHistoryIndex =>
            Session.SnapshotHistory.Count == 0 ? -1 : Session.SnapshotHistory[Session.SnapshotHistory.Count - 1].Index;

        public static SnapshotRecord GetSnapshotAt(int index)
        {
            foreach (SnapshotRecord record in Session.SnapshotHistory)
            {
                if (record.Index == index)
                {
                    return record;
                }
            }

            return null;
        }

        public static void ClearSnapshotHistory()
        {
            Session.SnapshotHistory.Clear();
            Session.NextSnapshotHistoryIndex = 0;
            System.Diagnostics.Debug.WriteLine("[SnapshotHistory] cleared");
        }

        public static ProcessDocumentCacheEntry LastProcessCache => Session.LastProcessCache;

        public static void SetLastProcessCache(ProcessDocumentCacheEntry entry)
        {
            Session.LastProcessCache = entry;
        }

        public static void ClearProcessDocumentCache(string reason = null)
        {
            foreach (DocumentSessionState session in SessionsByUuid.Values)
            {
                session.LastProcessCache = null;
            }

            Session.LastProcessCache = null;
            string reasonPart = string.IsNullOrEmpty(reason) ? "" : $" reason={reason}";
            System.Diagnostics.Debug.WriteLine($"[ProcessCache] cleared all sessions{reasonPart}");
        }

        public static Dictionary<string, string> NameToContentMap
        {
            get
            {
                var map = new Dictionary<string, string>();
                foreach (var mapping in Session.SentenceNameMapping)
                {
                    if (mapping.Count >= 3)
                    {
                        map[mapping[0]] = mapping[2];
                    }
                }

                return map;
            }
        }

        public static void SetSnapshots(string snapshot, string segSnapshot)
        {
            SetSnapshots(snapshot, segSnapshot, source: null);
        }

        public static void SetSnapshots(string snapshot, string segSnapshot, string source)
        {
            string snap = snapshot ?? "";
            string segSnap = segSnapshot ?? "";
            string docUuid = CurrDocUuid ?? "";

            var record = new SnapshotRecord
            {
                Index = Session.NextSnapshotHistoryIndex++,
                CreatedAtUtc = DateTime.UtcNow,
                Snapshot = snap,
                SegSnapshot = segSnap,
                Source = source ?? "",
                DocUuid = docUuid
            };
            Session.SnapshotHistory.Add(record);

            Session.Snapshot = snap;
            Session.SegSnapshot = segSnap;
            PrunePreReplaceCache(GetOrderedSentenceCodesFromSnapshot(Session.Snapshot));

            EasyWriteDiagnostics.LogDocumentExtractVerbose(
                $"[SnapshotHistory] append index={record.Index} source={record.Source} " +
                $"doc={docUuid} historyCount={Session.SnapshotHistory.Count} " +
                $"snapshotLen={snap.Length} segSnapshotLen={segSnap.Length}");
        }

        public static string CurrDocUuid => _activeDocUuid ?? "";

        public static void SetCurrDocUuid(string docUuid)
        {
            if (_activeSession == null)
            {
                return;
            }

            ApplyUuidToSession(_activeSession, docUuid, clearVolatileOnChange: true);
        }

        /// <summary>
        /// API 回写 uuid 时绑定到指定 Word 文档分片。
        /// </summary>
        public static void SetCurrDocUuidForDocument(Word.Document document, string docUuid)
        {
            BindAndActivate(document);
            ApplyUuidToSession(_activeSession, docUuid, clearVolatileOnChange: true);
            DocumentIdentity.UpdateUuid(document, docUuid);
        }

        private static void ApplyUuidToSession(DocumentSessionState session, string docUuid, bool clearVolatileOnChange)
        {
            if (session == null)
            {
                return;
            }

            string next = docUuid ?? "";
            string old = session.DocUuid ?? "";
            if (string.Equals(old, next, StringComparison.Ordinal))
            {
                session.DocUuid = next;
                _activeDocUuid = next;
                return;
            }

            if (clearVolatileOnChange)
            {
                session.ClearVolatileState();
            }

            session.DocUuid = next;
            if (!string.IsNullOrEmpty(old) && !string.Equals(old, next, StringComparison.Ordinal))
            {
                SessionsByUuid.Remove(old);
            }

            SessionsByUuid[next] = session;
            _activeDocUuid = next;
        }

        public static string EnsureCurrDocUuid(Word.Document document)
        {
            BindAndActivate(document);
            return DocumentIdentity.EnsureUuid(document);
        }

        /// <summary>
        /// 兼容旧调用：要求已有活动分片，否则返回空字符串。
        /// </summary>
        public static string EnsureCurrDocUuid()
        {
            if (!string.IsNullOrEmpty(CurrDocUuid))
            {
                return CurrDocUuid;
            }

            System.Diagnostics.Debug.WriteLine("[DocumentState] ⚠️ EnsureCurrDocUuid() 无活动文档绑定，请传入 Word.Document");
            return "";
        }

        public static void SetSentenceMappings(List<List<string>> sentenceMappings)
        {
            Session.SentenceNameMapping.Clear();
            if (sentenceMappings != null)
            {
                foreach (var mapping in sentenceMappings)
                {
                    if (mapping != null && mapping.Count >= 4)
                    {
                        Session.SentenceNameMapping.Add(new List<string> { mapping[0], mapping[1], mapping[2], mapping[3] });
                    }
                    else if (mapping != null && mapping.Count >= 3)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DocumentState] ⚠️ 警告：句子映射格式不正确，期望四元组，实际收到 {mapping.Count} 个元素，已跳过");
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置句子映射，共 {Session.SentenceNameMapping.Count} 条");
        }

        public static int SnapshotIdx => Session.SnapshotIdx;

        public static void SetSnapshotIdx(int snapshotIdx)
        {
            Session.SnapshotIdx = snapshotIdx;
            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置快照索引: {Session.SnapshotIdx}");
        }

        public static void SetSentenceNameToHeadingMap(Dictionary<string, Dictionary<string, object>> sentenceNameToHeadingMap)
        {
            Session.SentenceNameToHeadingMap.Clear();
            if (sentenceNameToHeadingMap != null)
            {
                foreach (var kvp in sentenceNameToHeadingMap)
                {
                    Dictionary<string, object> copiedValue = new Dictionary<string, object>();
                    foreach (var innerKvp in kvp.Value)
                    {
                        copiedValue[innerKvp.Key] = innerKvp.Value;
                    }

                    Session.SentenceNameToHeadingMap[kvp.Key] = copiedValue;
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置句子名称到标题信息映射，共 {Session.SentenceNameToHeadingMap.Count} 条");
        }

        public static Dictionary<string, Dictionary<string, object>> SentenceNameToHeadingMap
        {
            get
            {
                var map = new Dictionary<string, Dictionary<string, object>>();
                foreach (var kvp in Session.SentenceNameToHeadingMap)
                {
                    Dictionary<string, object> copiedValue = new Dictionary<string, object>();
                    foreach (var innerKvp in kvp.Value)
                    {
                        copiedValue[innerKvp.Key] = innerKvp.Value;
                    }

                    map[kvp.Key] = copiedValue;
                }

                return map;
            }
        }

        public static void SetSubtypeFormatMapping(List<Dictionary<string, object>> subtypeFormatMapping)
        {
            Session.SubtypeFormatMapping.Clear();
            if (subtypeFormatMapping != null)
            {
                foreach (var item in subtypeFormatMapping)
                {
                    Dictionary<string, object> copiedItem = CopyDictionary(item);
                    Session.SubtypeFormatMapping.Add(copiedItem);
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置detailed_subtype格式映射，共 {Session.SubtypeFormatMapping.Count} 条");
        }

        /// <summary>写入正文轻量众数；null 或空字典表示清空。不触碰 SubtypeFormatMapping。</summary>
        public static void SetBodyFormat(Dictionary<string, object> bodyFormat)
        {
            if (bodyFormat == null || bodyFormat.Count == 0)
            {
                Session.BodyFormat = null;
                System.Diagnostics.Debug.WriteLine("[DocumentState] BodyFormat 已清空");
                return;
            }

            Session.BodyFormat = CopyDictionary(bodyFormat);
            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置 BodyFormat，keys={Session.BodyFormat.Count}");
        }

        public static bool HasBodyFormat
        {
            get
            {
                return Session.BodyFormat != null && Session.BodyFormat.Count > 0;
            }
        }

        public static Dictionary<string, object> BodyFormat
        {
            get
            {
                if (Session.BodyFormat == null || Session.BodyFormat.Count == 0)
                {
                    return new Dictionary<string, object>();
                }

                return CopyDictionary(Session.BodyFormat);
            }
        }

        public static List<Dictionary<string, object>> SubtypeFormatMapping
        {
            get
            {
                var list = new List<Dictionary<string, object>>();
                foreach (var item in Session.SubtypeFormatMapping)
                {
                    list.Add(CopyDictionary(item));
                }

                return list;
            }
        }

        public static void SetRecognizedHeadingsInDocumentOrder(List<Dictionary<string, object>> headings)
        {
            Session.RecognizedHeadingsInDocumentOrder.Clear();
            if (headings == null)
            {
                return;
            }

            foreach (var item in headings)
            {
                Session.RecognizedHeadingsInDocumentOrder.Add(CopyDictionary(item));
            }

            System.Diagnostics.Debug.WriteLine($"[DocumentState] 已设置原文顺序标题映射，共 {Session.RecognizedHeadingsInDocumentOrder.Count} 条");
        }

        public static List<Dictionary<string, object>> RecognizedHeadingsInDocumentOrder
        {
            get
            {
                var list = new List<Dictionary<string, object>>();
                foreach (var item in Session.RecognizedHeadingsInDocumentOrder)
                {
                    list.Add(CopyDictionary(item));
                }

                return list;
            }
        }

        public static void AttachPreReplaceToNewCodes(IList<string> newCodes, Dictionary<string, object> snapshot)
        {
            if (newCodes == null || newCodes.Count == 0 || snapshot == null || snapshot.Count == 0)
            {
                return;
            }

            Dictionary<string, object> copy = CloneSnapshot(snapshot);
            foreach (string code in newCodes)
            {
                if (!string.IsNullOrEmpty(code) && code.StartsWith("S_", StringComparison.Ordinal))
                {
                    Session.PreReplaceByNewCode[code] = CloneSnapshot(copy);
                }
            }
        }

        public static Dictionary<string, object> TryGetPreReplaceFormat(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            if (Session.PreReplaceByNewCode.TryGetValue(code, out Dictionary<string, object> snapshot))
            {
                return CloneSnapshot(snapshot);
            }

            return null;
        }

        private static void PrunePreReplaceCache(IEnumerable<string> validCodes)
        {
            var valid = new HashSet<string>(validCodes ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
            var stale = Session.PreReplaceByNewCode.Keys.Where(k => !valid.Contains(k)).ToList();
            foreach (string key in stale)
            {
                Session.PreReplaceByNewCode.Remove(key);
            }
        }

        public static List<string> GetOrderedSentenceCodesFromSnapshot(string snapshot)
        {
            if (string.IsNullOrWhiteSpace(snapshot))
            {
                return new List<string>();
            }

            return snapshot
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s) && s.StartsWith("S_", StringComparison.Ordinal))
                .ToList();
        }

        public static List<string> GetOrderedSnapshotCodes()
        {
            return GetOrderedSentenceCodesFromSnapshot(Session.Snapshot);
        }

        public static List<int> GetSnapshotPositionsForCode(string code, string snapshot = null)
        {
            var positions = new List<int>();
            if (string.IsNullOrEmpty(code))
            {
                return positions;
            }

            string snap = snapshot ?? Session.Snapshot ?? "";
            var codes = GetOrderedSentenceCodesFromSnapshot(snap);
            for (int i = 0; i < codes.Count; i++)
            {
                if (string.Equals(codes[i], code, StringComparison.Ordinal))
                {
                    positions.Add(i);
                }
            }

            return positions;
        }

        public static int[] BuildSegIndexBySnapshotPosition(string snapshot = null, string segSnapshot = null)
        {
            var codes = GetOrderedSentenceCodesFromSnapshot(snapshot ?? Session.Snapshot);
            var segIndices = new int[codes.Count];
            if (codes.Count == 0)
            {
                return segIndices;
            }

            string segSnap = segSnapshot ?? Session.SegSnapshot ?? "";
            if (string.IsNullOrEmpty(segSnap))
            {
                return segIndices;
            }

            string[] segParts = segSnap.Split(';');
            int pos = 0;
            for (int segIndex = 0; segIndex < segParts.Length && pos < codes.Count; segIndex++)
            {
                string[] segCodes = segParts[segIndex]
                    .Split(',')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray();

                foreach (string segCode in segCodes)
                {
                    if (pos >= codes.Count)
                    {
                        break;
                    }

                    if (string.Equals(codes[pos], segCode, StringComparison.Ordinal))
                    {
                        segIndices[pos] = segIndex;
                        pos++;
                    }
                    else
                    {
                        pos++;
                    }
                }
            }

            return segIndices;
        }

        public static bool TryResolveOccurrence(
            string code,
            int occurrenceIndex,
            out int snapshotPosition,
            out int wordFindOrdinal,
            string snapshot = null)
        {
            snapshotPosition = -1;
            wordFindOrdinal = occurrenceIndex;

            if (occurrenceIndex < 0)
            {
                return false;
            }

            var positions = GetSnapshotPositionsForCode(code, snapshot);
            if (occurrenceIndex >= positions.Count)
            {
                return false;
            }

            snapshotPosition = positions[occurrenceIndex];
            return true;
        }

        private static Dictionary<string, object> CloneSnapshot(Dictionary<string, object> snapshot)
        {
            var copy = new Dictionary<string, object>();
            if (snapshot == null)
            {
                return copy;
            }

            foreach (var kvp in snapshot)
            {
                copy[kvp.Key] = kvp.Value;
            }

            return copy;
        }

        private static Dictionary<string, object> CopyDictionary(Dictionary<string, object> item)
        {
            Dictionary<string, object> copiedItem = new Dictionary<string, object>();
            if (item == null)
            {
                return copiedItem;
            }

            foreach (var kvp in item)
            {
                if (kvp.Value is Dictionary<string, object> nestedDict)
                {
                    Dictionary<string, object> copiedNestedDict = new Dictionary<string, object>();
                    foreach (var nestedKvp in nestedDict)
                    {
                        copiedNestedDict[nestedKvp.Key] = nestedKvp.Value;
                    }

                    copiedItem[kvp.Key] = copiedNestedDict;
                }
                else if (kvp.Value is List<object> nestedList)
                {
                    copiedItem[kvp.Key] = new List<object>(nestedList);
                }
                else
                {
                    copiedItem[kvp.Key] = kvp.Value;
                }
            }

            return copiedItem;
        }
    }
}
