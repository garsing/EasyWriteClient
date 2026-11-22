using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace WordAddIn1
{
    /// <summary>套用工具的 path：JSON 根对象就是原来的入参。</summary>
    public static class FormatFileArgs
    {
        public static readonly string[] CharacterConflicts =
        {
            "items", "target_codes", "target_paragraph_codes", "char_format", "text"
        };

        public static readonly string[] ParagraphConflicts =
        {
            "items", "target_paragraph_codes", "para_format"
        };

        public static bool TrySubstitute(
            Dictionary<string, object> args,
            string[] conflicts,
            out Dictionary<string, object> substituted,
            out string error)
        {
            substituted = args;
            error = null;
            if (args == null || !HasValue(args, "path"))
            {
                return true;
            }

            foreach (string key in conflicts)
            {
                if (HasValue(args, key))
                {
                    error = "路径和条目不能一起传";
                    return false;
                }
            }

            string raw = args["path"]?.ToString();
            if (!FilePathResolver.TryResolve(raw, out ResolvedFilePath resolved, out error))
            {
                return false;
            }

            FilePathIoResult read = FilePathResolver.ReadAsync(resolved).GetAwaiter().GetResult();
            if (read == null || !read.Success)
            {
                error = read?.Error ?? "读取文件失败";
                return false;
            }

            object parsed;
            try
            {
                parsed = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(read.Text ?? "");
            }
            catch (Exception ex)
            {
                error = "JSON 无法读取: " + ex.Message;
                return false;
            }

            if (!(parsed is Dictionary<string, object> fileArgs))
            {
                error = "JSON 根对象必须是工具入参";
                return false;
            }

            fileArgs.Remove("path");
            CopyIfMissing(args, fileArgs, "channel_id");
            CopyIfMissing(args, fileArgs, "conversation_id");
            CopyIfMissing(args, fileArgs, "tool_call_id");
            substituted = fileArgs;
            return true;
        }

        private static void CopyIfMissing(
            Dictionary<string, object> from,
            Dictionary<string, object> to,
            string key)
        {
            if (from != null && to != null && !to.ContainsKey(key) && from.TryGetValue(key, out object value))
            {
                to[key] = value;
            }
        }

        private static bool HasValue(Dictionary<string, object> args, string key)
        {
            if (args == null || string.IsNullOrEmpty(key) || !args.TryGetValue(key, out object value) || value == null)
            {
                return false;
            }

            if (value is string text)
            {
                return text.Length > 0;
            }

            return true;
        }
    }
}
