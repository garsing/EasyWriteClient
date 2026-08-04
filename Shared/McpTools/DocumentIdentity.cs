using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Word.Document 与 doc_uuid 的运行时绑定（进程内、随文档关闭解除）。
    /// </summary>
    public static class DocumentIdentity
    {
        private static readonly Dictionary<int, string> ObjectIdToUuid =
            new Dictionary<int, string>();

        private static readonly HashSet<int> CloseHandlerRegistered = new HashSet<int>();

        private static readonly object SyncRoot = new object();

        public static void EnsureCloseHandler(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            int objectId = GetObjectId(document);
            lock (SyncRoot)
            {
                if (CloseHandlerRegistered.Contains(objectId))
                {
                    return;
                }

                CloseHandlerRegistered.Add(objectId);
            }

            Word.Document docRef = document;
            var docEvents = (Word.DocumentEvents2_Event)docRef;
            docEvents.Close += () => HandleDocumentClose(docRef, objectId);
        }

        private static void HandleDocumentClose(Word.Document document, int objectId)
        {
            try
            {
                DocumentCheckpointService.ClearOnDocumentClose(document);
                DocumentState.ClearOnDocumentClose(document);
                System.Diagnostics.Debug.WriteLine($"[DocumentIdentity] Document Close event objectId={objectId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DocumentIdentity] Document Close event error: {ex.Message}");
            }
            finally
            {
                lock (SyncRoot)
                {
                    CloseHandlerRegistered.Remove(objectId);
                }
            }
        }

        public static int GetObjectId(Word.Document document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            return RuntimeHelpers.GetHashCode(document);
        }

        public static string TryResolveUuid(Word.Document document)
        {
            if (document == null)
            {
                return null;
            }

            int objectId = GetObjectId(document);
            if (ObjectIdToUuid.TryGetValue(objectId, out string uuid))
            {
                return uuid;
            }

            return null;
        }

        /// <summary>
        /// 为 document 确保 uuid；若尚未绑定则生成并登记。
        /// </summary>
        public static string EnsureUuid(Word.Document document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            string existing = TryResolveUuid(document);
            if (!string.IsNullOrEmpty(existing))
            {
                return existing;
            }

            string uuid = Uuid7Generator.GenerateUuid7();
            Register(document, uuid);
            DocumentState.GetOrCreateSession(uuid);
            System.Diagnostics.Debug.WriteLine($"[DocumentIdentity] bind doc objectId={GetObjectId(document)} uuid={uuid}");
            return uuid;
        }

        public static void Register(Word.Document document, string docUuid)
        {
            if (document == null || string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            ObjectIdToUuid[GetObjectId(document)] = docUuid;
        }

        /// <summary>
        /// API 回写 uuid 可能与本地 provisional uuid 不同，迁移绑定键。
        /// </summary>
        public static void UpdateUuid(Word.Document document, string docUuid)
        {
            if (document == null || string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            int objectId = GetObjectId(document);
            ObjectIdToUuid[objectId] = docUuid;
        }

        public static void Unregister(Word.Document document)
        {
            if (document == null)
            {
                return;
            }

            int objectId = GetObjectId(document);
            if (ObjectIdToUuid.TryGetValue(objectId, out string uuid))
            {
                ObjectIdToUuid.Remove(objectId);
                System.Diagnostics.Debug.WriteLine($"[DocumentIdentity] unbind doc objectId={objectId} uuid={uuid}");
            }
        }

        public static void ClearAll()
        {
            lock (SyncRoot)
            {
                ObjectIdToUuid.Clear();
                CloseHandlerRegistered.Clear();
            }
        }
    }
}
