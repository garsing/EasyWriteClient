using System;
using System.Collections.Generic;
using System.IO;

namespace WordAddIn1
{
    internal static class OpenDocumentPath
    {
        public static string TryGetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            string value = args[key].ToString()?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        public static bool ParseBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return defaultValue;
            }

            object raw = args[key];
            if (raw is bool b)
            {
                return b;
            }

            string s = Convert.ToString(raw)?.Trim();
            if (string.IsNullOrEmpty(s))
            {
                return defaultValue;
            }

            if (bool.TryParse(s, out bool parsed))
            {
                return parsed;
            }

            if (s == "1" || string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (s == "0" || string.Equals(s, "no", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return defaultValue;
        }

        public static string NormalizePath(string path)
        {
            try
            {
                return Path.GetFullPath(path).TrimEnd('\\', '/');
            }
            catch (Exception)
            {
                return path?.Trim() ?? "";
            }
        }

        public static bool TryResolve(string rawPath, bool createBlank, out string fullPath, out string error)
        {
            fullPath = null;
            error = null;

            string path = StripWrappingAsciiQuotes(rawPath?.Trim() ?? "");
            if (string.IsNullOrEmpty(path))
            {
                error = "必须提供 path 参数";
                return false;
            }

            var candidates = new List<string>();
            void AddCandidate(string p)
            {
                if (string.IsNullOrEmpty(p))
                {
                    return;
                }

                foreach (string existing in candidates)
                {
                    if (string.Equals(existing, p, StringComparison.Ordinal))
                    {
                        return;
                    }
                }

                candidates.Add(p);
            }

            AddCandidate(path);
            if (path.IndexOf('"') >= 0)
            {
                AddCandidate(ReplaceAsciiQuotesWithCurly(path));
                AddCandidate(path.Replace("\"", "\uFF02"));
            }

            Exception lastIllegal = null;
            string lastResolveError = null;
            foreach (string candidate in candidates)
            {
                if (!FilePathResolver.TryResolve(candidate, out ResolvedFilePath mapped, out string resolveError))
                {
                    lastResolveError = resolveError;
                    try
                    {
                        Path.GetFullPath(candidate);
                    }
                    catch (Exception ex)
                    {
                        lastIllegal = ex;
                    }

                    continue;
                }

                bool exists = File.Exists(mapped.LocalPath);
                if (createBlank)
                {
                    if (!exists)
                    {
                        fullPath = mapped.LocalPath;
                        return true;
                    }

                    continue;
                }

                if (exists)
                {
                    fullPath = mapped.LocalPath;
                    return true;
                }
            }

            if (createBlank)
            {
                foreach (string candidate in candidates)
                {
                    if (!FilePathResolver.TryResolve(candidate, out ResolvedFilePath mapped, out string resolveError))
                    {
                        lastResolveError = resolveError;
                        continue;
                    }

                    if (File.Exists(mapped.LocalPath))
                    {
                        error = "create_blank=true 但路径已存在，拒绝覆盖: " + mapped.LocalPath;
                        return false;
                    }

                    fullPath = mapped.LocalPath;
                    return true;
                }

                error = lastResolveError
                    ?? (lastIllegal != null
                        ? "路径非法: " + lastIllegal.Message
                          + "（文件名勿用英文引号 \"，请用中文 “” 或去掉引号）"
                        : "无法解析新建路径");
                return false;
            }

            if (!string.IsNullOrEmpty(lastResolveError) && candidates.Count == 1)
            {
                error = lastResolveError;
                return false;
            }

            if (lastIllegal != null && candidates.Count == 1)
            {
                error = "路径非法: " + lastIllegal.Message
                    + "（文件名勿用英文引号 \"，请用中文 “” 或从资源管理器复制真实路径）";
                return false;
            }

            error = "文件不存在（create_blank=false）: " + path
                + (path.IndexOf('"') >= 0
                    ? "。已尝试将英文引号替换为中文/全角引号仍未找到，请核对真实文件名"
                    : "");
            return false;
        }

        private static string StripWrappingAsciiQuotes(string path)
        {
            if (string.IsNullOrEmpty(path) || path.Length < 2)
            {
                return path;
            }

            if (path[0] == '"' && path[path.Length - 1] == '"')
            {
                return path.Substring(1, path.Length - 2).Trim();
            }

            return path;
        }

        private static string ReplaceAsciiQuotesWithCurly(string path)
        {
            if (string.IsNullOrEmpty(path) || path.IndexOf('"') < 0)
            {
                return path;
            }

            var chars = path.ToCharArray();
            bool left = true;
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] != '"')
                {
                    continue;
                }

                chars[i] = left ? '\u201C' : '\u201D';
                left = !left;
            }

            return new string(chars);
        }
    }
}
