using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 内部 WS document.writeCaptureImage：KB 截图落盘。不对模型注册。
    /// </summary>
    internal static class WriteCaptureImageHelper
    {
        public static async Task<ToolResult> ExecuteAsync(Dictionary<string, object> args)
        {
            string path = FilePathResolver.TryGetArg(args, "path");
            if (string.IsNullOrWhiteSpace(path))
            {
                return new ToolResult { Success = false, Error = "必须提供 path" };
            }

            string format = TryGetString(args, "format");
            byte[] bytes = TryGetImageBytes(args);
            if (bytes == null || bytes.Length == 0)
            {
                return new ToolResult { Success = false, Error = "必须提供 image_base64" };
            }

            var data = new Dictionary<string, object>();
            bool workspace = FilePathResolver.TryResolve(path, out ResolvedFilePath resolved, out _)
                && resolved.Kind == FilePathKind.Workspace;

            Func<Task<ToolResult>> write = async () =>
            {
                ToolResult writeError = await CapturePathWriter
                    .TryWriteAsync(path.Trim(), bytes, format, data)
                    .ConfigureAwait(false);
                if (writeError != null)
                {
                    return writeError;
                }

                return new ToolResult { Success = true, Data = data };
            };

            if (workspace)
            {
                return await WorkspaceReconcile
                    .AroundFrontendWriteAsync(WorkspaceReconcile.DefaultClientTtlSec, write)
                    .ConfigureAwait(false);
            }

            return await write().ConfigureAwait(false);
        }

        private static string TryGetString(IReadOnlyDictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }

        private static byte[] TryGetImageBytes(IReadOnlyDictionary<string, object> args)
        {
            if (args == null)
            {
                return null;
            }

            if (args.TryGetValue("image_base64", out object raw) && raw != null)
            {
                string b64 = Convert.ToString(raw)?.Trim();
                if (string.IsNullOrEmpty(b64))
                {
                    return null;
                }

                try
                {
                    return Convert.FromBase64String(b64);
                }
                catch (FormatException)
                {
                    return null;
                }
            }

            return null;
        }
    }
}
