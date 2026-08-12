using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WordAddIn1
{
    /// <summary>
    /// WPS Document RCW 与 doc_uuid 的运行时绑定（进程内；关闭时由 OpenFiles / 显式 Unregister 清理）。
    /// </summary>
    public static class WpsDocumentIdentity
    {
        private static readonly Dictionary<int, string> ObjectIdToUuid =
            new Dictionary<int, string>();

        private static readonly object SyncRoot = new object();

        public static int GetObjectId(object document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            return RuntimeHelpers.GetHashCode(document);
        }

        public static string TryResolveUuid(object document)
        {
            if (document == null)
            {
                return null;
            }

            int objectId = GetObjectId(document);
            lock (SyncRoot)
            {
                return ObjectIdToUuid.TryGetValue(objectId, out string uuid) ? uuid : null;
            }
        }

        public static string EnsureUuid(object document)
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
            System.Diagnostics.Debug.WriteLine(
                $"[WpsDocumentIdentity] bind doc objectId={GetObjectId(document)} uuid={uuid}");
            return uuid;
        }

        public static void Register(object document, string docUuid)
        {
            if (document == null || string.IsNullOrEmpty(docUuid))
            {
                return;
            }

            lock (SyncRoot)
            {
                ObjectIdToUuid[GetObjectId(document)] = docUuid;
            }
        }

        public static void Unregister(object document)
        {
            if (document == null)
            {
                return;
            }

            int objectId = GetObjectId(document);
            lock (SyncRoot)
            {
                if (ObjectIdToUuid.TryGetValue(objectId, out string uuid))
                {
                    ObjectIdToUuid.Remove(objectId);
                    System.Diagnostics.Debug.WriteLine(
                        $"[WpsDocumentIdentity] unbind doc objectId={objectId} uuid={uuid}");
                }
            }
        }

        /// <summary>关文档：清 session + 渠道 + 身份绑定。</summary>
        public static void ClearOnDocumentClose(object document)
        {
            if (document == null)
            {
                return;
            }

            string uuid = TryResolveUuid(document);
            Unregister(document);
            if (!string.IsNullOrEmpty(uuid))
            {
                DocumentState.ClearSessionByUuid(uuid);
            }
        }

        public static void ClearAll()
        {
            lock (SyncRoot)
            {
                ObjectIdToUuid.Clear();
            }
        }
    }
}
