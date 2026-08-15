using System;
using System.IO;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetContentUtil
    {
        public static bool ShouldSkipWorkbook(string name, string fullName, bool isAddin)
        {
            if (isAddin)
            {
                return true;
            }

            string file = name;
            if (string.IsNullOrEmpty(file) && !string.IsNullOrEmpty(fullName))
            {
                try
                {
                    file = Path.GetFileName(fullName);
                }
                catch (Exception)
                {
                    file = fullName;
                }
            }

            return !string.IsNullOrEmpty(file)
                && string.Equals(file, "PERSONAL.XLSB", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsSavedPath(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return false;
            }

            if (fullName.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return fullName.IndexOf(Path.DirectorySeparatorChar) >= 0
                || fullName.IndexOf(Path.AltDirectorySeparatorChar) >= 0;
        }

        public static string NormalizePath(string path)
        {
            try
            {
                return Path.GetFullPath(path).TrimEnd('\\', '/');
            }
            catch (Exception)
            {
                return path?.Trim();
            }
        }

        public static bool IsHiddenVisibleState(object visible)
        {
            if (visible == null)
            {
                return false;
            }

            try
            {
                int value = Convert.ToInt32(visible);
                return value != -1;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
