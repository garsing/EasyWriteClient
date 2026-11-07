using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace WordAddIn1
{
    /// <summary>
    /// 文档自定义属性 EasyWriteDocId。副本目录用这个值，不用内存里的文档 uuid。
    /// </summary>
    public static class EasyWriteDocIdStore
    {
        public const string PropertyName = "EasyWriteDocId";

        private const int MsoPropertyTypeString = 4;

        [ThreadStatic]
        private static bool _savingSplit;

        public static string TryRead(object live)
        {
            object props = TryGetCustomProperties(live);
            if (props == null)
            {
                return null;
            }

            try
            {
                foreach (object prop in (IEnumerable)props)
                {
                    if (prop == null)
                    {
                        continue;
                    }

                    string name = Convert.ToString(GetMember(prop, "Name"), CultureInfo.InvariantCulture);
                    if (!string.Equals(name, PropertyName, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    string value = Convert.ToString(GetMember(prop, "Value"), CultureInfo.InvariantCulture);
                    return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[EasyWriteDocId] read failed: " + ex.Message);
            }

            return null;
        }

        public static bool TryWrite(object live, string docId)
        {
            if (live == null || string.IsNullOrWhiteSpace(docId))
            {
                return false;
            }

            object props = TryGetCustomProperties(live);
            if (props == null)
            {
                return false;
            }

            try
            {
                foreach (object prop in (IEnumerable)props)
                {
                    if (prop == null)
                    {
                        continue;
                    }

                    string name = Convert.ToString(GetMember(prop, "Name"), CultureInfo.InvariantCulture);
                    if (!string.Equals(name, PropertyName, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    SetMember(prop, "Value", docId.Trim());
                    return true;
                }

                props.GetType().InvokeMember(
                    "Add",
                    BindingFlags.InvokeMethod,
                    null,
                    props,
                    new object[] { PropertyName, false, MsoPropertyTypeString, docId.Trim() });
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[EasyWriteDocId] write failed: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 另存为之后活文档路径和这份身份记下的 path 不同，就换新号并再保存一次。
        /// </summary>
        public static string TrySplitOnSaveAs(object live)
        {
            if (live == null || _savingSplit)
            {
                return TryRead(live);
            }

            string currentId = TryRead(live);
            if (string.IsNullOrEmpty(currentId))
            {
                return null;
            }

            string fullName = TryReadFullName(live);
            if (!IsPersistedPath(fullName))
            {
                return currentId;
            }

            DocumentCheckpointSession session = DocumentCheckpointService.PeekSession(currentId);
            string recorded = session?.Path;
            if (string.IsNullOrEmpty(recorded)
                || string.Equals(NormalizePath(recorded), NormalizePath(fullName), StringComparison.OrdinalIgnoreCase))
            {
                return currentId;
            }

            string newId = Uuid7Generator.GenerateUuid7();
            if (!TryWrite(live, newId))
            {
                return currentId;
            }

            DocumentCheckpointService.StartFreshIdentity(newId, fullName, live);
            try
            {
                _savingSplit = true;
                live.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, live, null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[EasyWriteDocId] save after split failed: " + ex.Message);
            }
            finally
            {
                _savingSplit = false;
            }

            return newId;
        }

        public static string TryReadFullName(object live)
        {
            if (live == null)
            {
                return "";
            }

            try
            {
                return Convert.ToString(GetMember(live, "FullName"), CultureInfo.InvariantCulture) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        public static string TryReadName(object live)
        {
            if (live == null)
            {
                return "";
            }

            try
            {
                string name = Convert.ToString(GetMember(live, "Name"), CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name.Trim();
                }
            }
            catch (Exception)
            {
            }

            string full = TryReadFullName(live);
            return string.IsNullOrEmpty(full) ? "" : Path.GetFileName(full);
        }

        public static bool IsPersistedPath(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return false;
            }

            if (fullName.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return fullName.IndexOf('\\') >= 0
                || fullName.IndexOf('/') >= 0
                || File.Exists(fullName);
        }

        public static string KindFromExtension(string extension)
        {
            if (string.Equals(extension, "xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return "excel";
            }

            if (string.Equals(extension, "pptx", StringComparison.OrdinalIgnoreCase))
            {
                return "ppt";
            }

            return "word";
        }

        private static string NormalizePath(string path)
        {
            return (path ?? "").Replace('/', '\\').Trim();
        }

        private static object TryGetCustomProperties(object live)
        {
            if (live == null)
            {
                return null;
            }

            try
            {
                return GetMember(live, "CustomDocumentProperties");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[EasyWriteDocId] properties failed: " + ex.Message);
                return null;
            }
        }

        private static object GetMember(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                null);
        }

        private static void SetMember(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                new[] { value });
        }
    }
}
