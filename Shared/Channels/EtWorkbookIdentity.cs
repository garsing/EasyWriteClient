using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WordAddIn1
{
    public static class EtWorkbookIdentity
    {
        private static readonly Dictionary<int, string> ObjectIdToUuid = new Dictionary<int, string>();
        private static readonly object SyncRoot = new object();

        public static int GetObjectId(object workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            return RuntimeHelpers.GetHashCode(workbook);
        }

        public static string TryResolveUuid(object workbook)
        {
            if (workbook == null)
            {
                return null;
            }

            lock (SyncRoot)
            {
                return ObjectIdToUuid.TryGetValue(GetObjectId(workbook), out string uuid) ? uuid : null;
            }
        }

        public static string EnsureUuid(object workbook)
        {
            string existing = TryResolveUuid(workbook);
            if (!string.IsNullOrEmpty(existing))
            {
                return existing;
            }

            string uuid = Uuid7Generator.GenerateUuid7();
            lock (SyncRoot)
            {
                ObjectIdToUuid[GetObjectId(workbook)] = uuid;
            }

            return uuid;
        }

        public static void Unregister(object workbook)
        {
            if (workbook == null)
            {
                return;
            }

            lock (SyncRoot)
            {
                ObjectIdToUuid.Remove(GetObjectId(workbook));
            }
        }
    }
}
