using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WordAddIn1
{
    public static class WppPresentationIdentity
    {
        private static readonly Dictionary<int, string> ObjectIdToUuid = new Dictionary<int, string>();
        private static readonly object SyncRoot = new object();

        public static int GetObjectId(object presentation)
        {
            if (presentation == null)
            {
                throw new ArgumentNullException(nameof(presentation));
            }

            return RuntimeHelpers.GetHashCode(presentation);
        }

        public static string TryResolveUuid(object presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            lock (SyncRoot)
            {
                return ObjectIdToUuid.TryGetValue(GetObjectId(presentation), out string uuid) ? uuid : null;
            }
        }

        public static string EnsureUuid(object presentation)
        {
            string existing = TryResolveUuid(presentation);
            if (!string.IsNullOrEmpty(existing))
            {
                return existing;
            }

            string uuid = Uuid7Generator.GenerateUuid7();
            lock (SyncRoot)
            {
                ObjectIdToUuid[GetObjectId(presentation)] = uuid;
            }

            return uuid;
        }

        public static void Unregister(object presentation)
        {
            if (presentation == null)
            {
                return;
            }

            lock (SyncRoot)
            {
                ObjectIdToUuid.Remove(GetObjectId(presentation));
            }
        }
    }
}
