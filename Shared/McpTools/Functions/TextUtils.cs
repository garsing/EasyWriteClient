using System;

namespace WordAddIn1
{
    /// <summary>
    /// 文本工具模块
    /// 提供文本转义和反转义功能，用于数据库存储、调试和打印输出
    /// </summary>
    public static class TextUtils
    {
        /// <summary>
        /// 将换行符转义为可见字符串，用于数据库存储和调试
        /// 
        /// 转义规则：
        /// - \r\n -> \\r\\n
        /// - \r -> \\r
        /// - \n -> \\n
        /// 
        /// 注意：先转义 \r\n，再转义 \r 和 \n，避免重复转义
        /// </summary>
        /// <param name="content">原始内容</param>
        /// <returns>转义后的内容</returns>
        public static string EscapeNewlinesForDb(string content)
        {
            if (content == null)
            {
                return null;
            }

            // 先转义 \r\n，再转义 \r 和 \n
            string escaped = content.Replace("\r\n", "\\r\\n");
            escaped = escaped.Replace("\r", "\\r");
            escaped = escaped.Replace("\n", "\\n");

            return escaped;
        }

        /// <summary>
        /// 将数据库中的转义字符串反转义为原始换行符
        /// 
        /// 反转义规则：
        /// - \\r\\n -> \r\n
        /// - \\r -> \r
        /// - \\n -> \n
        /// 
        /// 注意：先反转义 \\r\\n，再反转义 \\r 和 \\n，避免重复反转义
        /// </summary>
        /// <param name="content">转义后的内容（从数据库读取）</param>
        /// <returns>反转义后的原始内容</returns>
        public static string UnescapeNewlinesFromDb(string content)
        {
            if (content == null)
            {
                return null;
            }

            // 先反转义 \\r\\n，再反转义 \\r 和 \\n
            string unescaped = content.Replace("\\r\\n", "\r\n");
            unescaped = unescaped.Replace("\\r", "\r");
            unescaped = unescaped.Replace("\\n", "\n");

            return unescaped;
        }

        /// <summary>
        /// 切句与映射 hash 前的换行归一化。
        /// Readback：保留 WordReader 的 \n（w:br 手动换行），仅合并 \r\n。
        /// ActionContent：大模型写入参数，将 \n 也归一化为 \r。
        /// </summary>
        public static string PrepareForSentenceSplit(string text, SentenceSplitNewlineMode mode)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text ?? "";
            }

            string normalized = text.Replace("\r\n", "\r");
            if (mode == SentenceSplitNewlineMode.ActionContent)
            {
                normalized = normalized.Replace("\n", "\r");
            }

            return normalized;
        }

        /// <summary>
        /// Replace 新内容段尾对齐：在大模型输入后、分句/编码处理前调用。
        /// 原文末尾无段落符 → 不改动；原文末尾有段落符 → 去掉新文段尾所有段落符后，
        /// 附上原文段尾后缀（数量与字符序列均与原文一致），便于后续逐段映射段落类型。
        /// </summary>
        public static string EnsureReplaceContentTrailingParagraphMarks(string newContent, string oldText)
        {
            if (string.IsNullOrEmpty(newContent))
            {
                return newContent ?? "";
            }

            string oldSuffix = GetTrailingParagraphSuffix(oldText);
            if (string.IsNullOrEmpty(oldSuffix))
            {
                return newContent;
            }

            return TrimTrailingParagraphMarks(newContent) + oldSuffix;
        }

        /// <summary>去掉文本末尾连续的段落符（\r、\n）。</summary>
        public static string TrimTrailingParagraphMarks(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text ?? "";
            }

            int end = text.Length;
            while (end > 0)
            {
                char c = text[end - 1];
                if (c == '\r' || c == '\n')
                {
                    end--;
                }
                else
                {
                    break;
                }
            }

            return end < text.Length ? text.Substring(0, end) : text;
        }

        /// <summary>段尾连续段落符的数量（\r、\n 各计 1）。</summary>
        public static int CountTrailingParagraphMarks(string text)
        {
            return GetTrailingParagraphSuffix(text).Length;
        }

        /// <summary>提取文本末尾连续的段落符（\r、\n）。</summary>
        public static string GetTrailingParagraphSuffix(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            int i = text.Length;
            while (i > 0)
            {
                char c = text[i - 1];
                if (c == '\r' || c == '\n')
                {
                    i--;
                }
                else
                {
                    break;
                }
            }

            return i < text.Length ? text.Substring(i) : "";
        }

        /// <summary>文本是否以段落符（\r 或 \n）结尾。</summary>
        public static bool EndsWithParagraphMark(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            char last = text[text.Length - 1];
            return last == '\r' || last == '\n';
        }
    }
}

