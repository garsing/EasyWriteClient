using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 截图可选 path：FilePathResolver + 扩展名必须与压缩后 format 一致。
    /// </summary>
    internal static class CapturePathWriter
    {
        public static bool HasPath(Dictionary<string, object> args)
        {
            return !string.IsNullOrWhiteSpace(FilePathResolver.TryGetArg(args, "path"));
        }

        public static async Task<ToolResult> TryAttachPathAsync(
            Dictionary<string, object> args,
            byte[] bytes,
            string format,
            Dictionary<string, object> data)
        {
            string raw = FilePathResolver.TryGetArg(args, "path");
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            return await TryWriteAsync(raw.Trim(), bytes, format, data).ConfigureAwait(false);
        }

        public static async Task<ToolResult> TryWriteAsync(
            string rawPath,
            byte[] bytes,
            string format,
            Dictionary<string, object> data)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return Fail("截图字节为空，无法写入 path");
            }

            if (!FilePathResolver.TryResolve(rawPath, out ResolvedFilePath resolved, out string resolveError))
            {
                return Fail(resolveError);
            }

            string ext = Path.GetExtension(resolved.LocalPath) ?? "";
            string fmt = (format ?? "").Trim().ToLowerInvariant();
            if (string.Equals(fmt, "jpg", StringComparison.OrdinalIgnoreCase))
            {
                fmt = "jpeg";
            }

            if (!IsAllowedExtension(ext))
            {
                return Fail("path 扩展名须为 .png / .jpg / .jpeg");
            }

            if (!ExtensionMatchesFormat(ext, fmt))
            {
                return Fail("path 扩展名须与实际 format 一致（当前 format=" + fmt + "）");
            }

            var written = await FilePathResolver.WriteBytesAsync(resolved, bytes).ConfigureAwait(false);
            if (!written.Success)
            {
                return Fail(written.Error ?? "写入截图失败");
            }

            if (data != null)
            {
                data["path"] = resolved.Display ?? rawPath;
            }

            return null;
        }

        private static bool IsAllowedExtension(string ext)
        {
            return string.Equals(ext, ".png", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".jpg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".jpeg", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ExtensionMatchesFormat(string ext, string format)
        {
            if (string.Equals(format, "png", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(ext, ".png", StringComparison.OrdinalIgnoreCase);
            }

            if (string.Equals(format, "jpeg", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(ext, ".jpg", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(ext, ".jpeg", StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
