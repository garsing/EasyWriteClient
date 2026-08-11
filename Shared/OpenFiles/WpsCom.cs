using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// WPS 文字 COM 晚绑定辅助：仅 GetActiveObject，不 new / 不 Quit。
    /// </summary>
    internal static class WpsCom
    {
        public static readonly string[] ProgIds =
        {
            "Kwps.Application",
            "kwps.Application",
            "wps.Application"
        };

        public static object TryGetActiveApplication(out string progId)
        {
            progId = null;
            foreach (string id in ProgIds)
            {
                try
                {
                    object app = Marshal.GetActiveObject(id);
                    if (app != null && CanReadDocuments(app))
                    {
                        progId = id;
                        return app;
                    }
                }
                catch (COMException)
                {
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        public static bool CanReadDocuments(object app)
        {
            try
            {
                object docs = GetProperty(app, "Documents");
                if (docs == null)
                {
                    return false;
                }

                object countObj = GetProperty(docs, "Count");
                if (countObj == null)
                {
                    return false;
                }

                Convert.ToInt32(countObj);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsAlive(object app)
        {
            try
            {
                object name = GetProperty(app, "Name");
                return name != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static IEnumerable<object> EnumerateDocuments(object app)
        {
            if (app == null)
            {
                yield break;
            }

            object docs;
            int count;
            try
            {
                docs = GetProperty(app, "Documents");
                count = Convert.ToInt32(GetProperty(docs, "Count"));
            }
            catch (Exception)
            {
                yield break;
            }

            // Word/WPS 集合多为 1-based
            for (int i = 1; i <= count; i++)
            {
                object doc = null;
                try
                {
                    doc = GetIndexed(docs, i);
                }
                catch (Exception)
                {
                    doc = null;
                }

                if (doc != null)
                {
                    yield return doc;
                }
            }
        }

        public static string TryReadFullName(object doc)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                object full = GetProperty(doc, "FullName");
                string fullName = full as string;
                if (string.IsNullOrWhiteSpace(fullName))
                {
                    return null;
                }

                if (fullName.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                return fullName;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string TryReadName(object doc)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                return GetProperty(doc, "Name") as string;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static object GetProperty(object target, string name)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return null;
            }

            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                null);
        }

        private static object GetIndexed(object collection, int index)
        {
            // 优先 Item 属性/方法
            try
            {
                return collection.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    collection,
                    new object[] { index });
            }
            catch (Exception)
            {
            }

            try
            {
                return collection.GetType().InvokeMember(
                    "get_Item",
                    BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    collection,
                    new object[] { index });
            }
            catch (Exception)
            {
            }

            // 部分实现支持默认索引器
            return collection.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase,
                null,
                collection,
                new object[] { index });
        }
    }
}
