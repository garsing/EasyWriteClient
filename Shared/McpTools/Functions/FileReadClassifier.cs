using System;
using System.Collections.Generic;
using System.IO;

namespace WordAddIn1
{
    internal enum FileReadKind
    {
        Image,
        Text,
        RejectedBinary
    }

    /// <summary>
    /// F_read_file 按后缀分流（读文件图片视觉轮 I3 / I4）。
    /// </summary>
    internal static class FileReadClassifier
    {
        private static readonly HashSet<string> ImageExtensions = NewSet(
            "png", "jpg", "jpeg", "webp", "bmp", "gif");

        private static readonly HashSet<string> TextExtensions = NewSet(
            "txt", "csv", "tsv", "json", "yaml", "yml", "xml", "py", "md",
            "html", "htm", "css", "js", "log", "ini", "conf", "svg", "sql");

        private static readonly HashSet<string> RejectedBinaryExtensions = NewSet(
            "docx", "doc", "xlsx", "xls", "xlsm", "xlsb", "pptx", "ppt",
            "pdf", "exe", "dll", "msi", "zip", "7z", "rar", "gz",
            "heic", "heif", "avif", "tiff", "tif");

        public static string ExtensionOf(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return "";
            }

            string ext = Path.GetExtension(path.Trim());
            if (string.IsNullOrEmpty(ext))
            {
                return "";
            }

            return ext.TrimStart('.').ToLowerInvariant();
        }

        public static FileReadKind Classify(string path, byte[] bytes)
        {
            string ext = ExtensionOf(path);
            if (ImageExtensions.Contains(ext))
            {
                return FileReadKind.Image;
            }

            if (RejectedBinaryExtensions.Contains(ext))
            {
                return FileReadKind.RejectedBinary;
            }

            if (TextExtensions.Contains(ext))
            {
                return FileReadKind.Text;
            }

            return HasNulPrefix(bytes, 8192)
                ? FileReadKind.RejectedBinary
                : FileReadKind.Text;
        }

        private static bool HasNulPrefix(byte[] bytes, int maxBytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return false;
            }

            int n = Math.Min(bytes.Length, maxBytes);
            for (int i = 0; i < n; i++)
            {
                if (bytes[i] == 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static HashSet<string> NewSet(params string[] items)
        {
            return new HashSet<string>(items, StringComparer.OrdinalIgnoreCase);
        }
    }
}
