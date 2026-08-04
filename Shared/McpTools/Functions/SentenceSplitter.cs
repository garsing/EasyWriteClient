using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordAddIn1
{
    /// <summary>
    /// 分句前换行归一化模式。
    /// </summary>
    public enum SentenceSplitNewlineMode
    {
        /// <summary>Word 读回 readText：保留 \n（w:br），仅 \r\n → \r。</summary>
        Readback = 0,

        /// <summary>大模型 replace/insert 参数：\r\n、\n → \r。</summary>
        ActionContent = 1,
    }

    /// <summary>
    /// 句子切分模块
    /// 提供统一的句子切分函数，用于将文本按照句子结束符和换行符进行分割
    /// </summary>
    public static class SentenceSplitter
    {
        /// <summary>
        /// 文档切句：先按 <paramref name="newlineMode"/> 归一化换行，再分句。
        /// </summary>
        public static List<string> SplitTextBySentencesForDocument(
            string text,
            SentenceSplitNewlineMode newlineMode = SentenceSplitNewlineMode.Readback)
        {
            return SplitTextBySentences(TextUtils.PrepareForSentenceSplit(text, newlineMode));
        }

        /// <summary>Word 读回 seg 切句（保留 \n）。</summary>
        public static List<string> SplitForReadback(string text)
        {
            return SplitTextBySentencesForDocument(text, SentenceSplitNewlineMode.Readback);
        }

        /// <summary>大模型写入内容切句（\n → \r）。</summary>
        public static List<string> SplitForActionContent(string text)
        {
            return SplitTextBySentencesForDocument(text, SentenceSplitNewlineMode.ActionContent);
        }

        /// <summary>
        /// 根据句子结束符和换行符对文本进行分句
        /// 
        /// 分割优先级：
        /// 1. 优先使用句子结束符（。；？！：.;?!:）+ 换行符（\r\n、\r、\n）的组合来分割
        /// 2. 如果没有找到组合，则使用单独的句子结束符、\r、\n、\r\n、&lt;/table&gt;来分割
        /// 
        /// 【边界处理】以下情况中的英文点号不作为句子结束符：
        /// - 小数：3.14, 0.5（数字.数字）
        /// - 序号：1. 2. (1). (a). A. B. C.（数字/字母 + 点 + 空格/换行）
        /// - 常见缩写：Mr. Mrs. Dr. Prof. etc. i.e. e.g. U.S. U.K.
        /// - 文件后缀：.xlsm .doc .pdf .txt .xlsx .pptx（点 + 字母组合）
        /// 
        /// 算法逻辑：
        /// 1. 优先查找句子结束符+换行符的组合（如。\r\n、。\r、。\n）
        /// 2. 如果没有找到组合，则查找单独的句子结束符、\r、\n、\r\n、&lt;/table&gt;
        /// 3. 遇到英文点号(.)时，通过向前/向后查看上下文判断是否属于上述边界情况
        /// 4. 提取句子时，完全保留原文本，不做任何处理（不处理空格和制表符）
        /// 5. 如果句子不为空，添加到列表中
        /// 
        /// 注意：
        /// - 保持原文本原汁原味，不做任何处理（不统一换行符，不处理空格和制表符）
        /// - 换行符、空格、制表符等所有字符都会被保留在句子内容中
        /// - 不修改原文，仅通过逻辑判断是否切分
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <returns>句子列表（保持原样，包括所有字符）</returns>
        public static List<string> SplitTextBySentences(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new List<string>();
            }

            var sentences = new List<string>();
            var sentenceEndings = new[] { '。', '；', '？', '！', '：', '.', ';', '?', '!', ':' };

            int start = 0;
            int i = 0;

            while (i < text.Length)
            {
                bool foundSplit = false;
                int sentenceEnd = 0;

                // 【关键修改】对于英文点号，先检查是否属于不应分割的情况
                if (text[i] == '.')
                {
                    if (IsDecimalPoint(text, i) || IsOrdinalNumber(text, i) || 
                        IsAbbreviation(text, i) || IsLetterOrdinal(text, i) || 
                        IsFileExtension(text, i))
                    {
                        i++;
                        continue;
                    }
                }

                // 优先查找：句子结束符 + 换行符的组合
                if (sentenceEndings.Contains(text[i]))
                {
                    // 检查后面是否有换行符
                    if (i + 1 < text.Length)
                    {
                        if (text[i + 1] == '\r')
                        {
                            // 可能是 \r 或 \r\n
                            if (i + 2 < text.Length && text[i + 2] == '\n')
                            {
                                // 句子结束符 + \r\n 组合
                                sentenceEnd = i + 3;
                                foundSplit = true;
                            }
                            else
                            {
                                // 句子结束符 + \r 组合
                                sentenceEnd = i + 2;
                                foundSplit = true;
                            }
                        }
                        else if (text[i + 1] == '\n')
                        {
                            // 句子结束符 + \n 组合
                            sentenceEnd = i + 2;
                            foundSplit = true;
                        }
                    }

                    // 如果句子结束符后面没有换行符，也进行分割（"实在没有"的情况）
                    if (!foundSplit)
                    {
                        sentenceEnd = i + 1;
                        foundSplit = true;
                    }
                }

                // 如果没有找到句子结束符+换行符的组合，查找单独的换行符或 </table>
                if (!foundSplit)
                {
                    if (text[i] == '\r')
                    {
                        // 可能是 \r 或 \r\n
                        if (i + 1 < text.Length && text[i + 1] == '\n')
                        {
                            // \r\n 组合
                            sentenceEnd = i + 2;
                            foundSplit = true;
                        }
                        else
                        {
                            // 单独的 \r
                            sentenceEnd = i + 1;
                            foundSplit = true;
                        }
                    }
                    else if (text[i] == '\n')
                    {
                        // 单独的 \n
                        sentenceEnd = i + 1;
                        foundSplit = true;
                    }
                    else if (i + 8 <= text.Length && text.Substring(i, 8) == "</table>")
                    {
                        // </table> 标签，也作为分割符
                        sentenceEnd = i + 8;
                        foundSplit = true;
                    }
                }

                if (foundSplit)
                {
                    // 提取句子（包含结束符和换行符）
                    // 完全保留原文本，不做任何处理
                    string sentence = text.Substring(start, sentenceEnd - start);

                    // 如果句子只包含空白字符（空格、\n、\r、\t等），拼接到前一句中
                    if (IsWhitespaceOnly(sentence))
                    {
                        if (sentences.Count > 0)
                        {
                            // 拼接到前一句中
                            sentences[sentences.Count - 1] += sentence;
                        }
                        else
                        {
                            // 如果是第一个句子且只包含空白字符，仍然添加（保持完整性）
                            sentences.Add(sentence);
                        }
                    }
                    else
                    {
                        // 正常添加句子
                        sentences.Add(sentence);
                    }

                    start = sentenceEnd;
                    i = sentenceEnd; // 从分割位置之后继续查找
                }
                else
                {
                    // 不是分割符，继续查找
                    i++;
                }
            }

            // 如果文本没有以分割符结尾，添加剩余部分
            if (start < text.Length)
            {
                string remaining = text.Substring(start);
                // 如果剩余部分只包含空白字符，拼接到前一句中
                if (IsWhitespaceOnly(remaining))
                {
                    if (sentences.Count > 0)
                    {
                        sentences[sentences.Count - 1] += remaining;
                    }
                    else
                    {
                        sentences.Add(remaining);
                    }
                }
                else
                {
                    sentences.Add(remaining);
                }
            }

            // 【关键验证】确保所有句子拼接后完全等于原文（不增不减）
            string reconstructed = string.Join("", sentences);
            if (reconstructed != text)
            {
                // 如果重构的文本与原文不一致，返回原文本作为单个句子
                // 这不应该发生，但如果发生，至少保证不丢失原文
                return new List<string> { text };
            }

            // 注意：保留句子末尾的 \r，以便在搜索时能够正确匹配跨段落的内容
            // 之前会去掉末尾的 \r，但现在保留以便支持跨段落搜索

            // 如果原文本为空，返回空列表；否则返回句子列表
            return string.IsNullOrEmpty(text) ? new List<string>() : sentences;
        }

        /// <summary>
        /// 检查是否是小数点（前后都是数字）
        /// </summary>
        private static bool IsDecimalPoint(string text, int pos)
        {
            if (pos >= text.Length || text[pos] != '.')
            {
                return false;
            }
            return (pos > 0 && pos < text.Length - 1 &&
                    char.IsDigit(text[pos - 1]) && char.IsDigit(text[pos + 1]));
        }

        /// <summary>
        /// 检查是否是序号中的点（如 1. (1). a.）
        /// </summary>
        private static bool IsOrdinalNumber(string text, int pos)
        {
            if (pos >= text.Length || text[pos] != '.')
            {
                return false;
            }
            if (pos > 0)
            {
                char charBefore = text[pos - 1];
                // 数字后直接跟点：1. 
                if (char.IsDigit(charBefore))
                {
                    return true;
                }
                // 括号/方括号内的数字/字母后跟点：(1). [a]. 中的点
                if ((charBefore == ')' || charBefore == ']') && pos > 1 && char.IsLetterOrDigit(text[pos - 2]))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 检查是否是字母序号中的点（如 A. B. C. a. b.）
        /// </summary>
        private static bool IsLetterOrdinal(string text, int pos)
        {
            if (pos >= text.Length || text[pos] != '.')
            {
                return false;
            }
            if (pos > 0)
            {
                char charBefore = text[pos - 1];
                // 单个字母（大小写）后直接跟点：A. B. a. b.
                if (char.IsLetter(charBefore))
                {
                    // 检查前面是否也是字母（排除单词中的点）
                    if (pos == 1) // 文本开头
                    {
                        // 检查后面是否是空格、换行符或结束
                        if (pos + 1 >= text.Length)
                        {
                            return true; // 文本末尾
                        }
                        char charAfter = text[pos + 1];
                        if (charAfter == ' ' || charAfter == '\t' || charAfter == '\r' || charAfter == '\n')
                        {
                            return true;
                        }
                        // 检查是否是下一个字母序号（如 A. B. C.）
                        if (char.IsLetter(charAfter) && (pos + 2 >= text.Length || text[pos + 2] == '.'))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        char charBeforePrev = text[pos - 2];
                        // 如果前面是空格、换行符、开始或标点符号，可能是字母序号
                        if (charBeforePrev == ' ' || charBeforePrev == '\t' || charBeforePrev == '\r' || 
                            charBeforePrev == '\n' || charBeforePrev == '(' || charBeforePrev == '[' || 
                            charBeforePrev == '{' || charBeforePrev == ',' || charBeforePrev == ';' || 
                            charBeforePrev == ':' || charBeforePrev == '!' || charBeforePrev == '?' || 
                            charBeforePrev == '.')
                        {
                            // 检查后面是否是空格、换行符或结束
                            if (pos + 1 >= text.Length)
                            {
                                return true; // 文本末尾
                            }
                            char charAfter = text[pos + 1];
                            if (charAfter == ' ' || charAfter == '\t' || charAfter == '\r' || charAfter == '\n')
                            {
                                return true;
                            }
                            // 检查是否是下一个字母序号（如 A. B. C.）
                            if (char.IsLetter(charAfter) && (pos + 2 >= text.Length || text[pos + 2] == '.'))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 检查是否是文件后缀中的点（如 .xlsm .doc .pdf .txt .xlsx .pptx）
        /// </summary>
        private static bool IsFileExtension(string text, int pos)
        {
            if (pos >= text.Length || text[pos] != '.')
            {
                return false;
            }

            // 文件后缀通常是：点 + 2-5个字母/数字
            // 检查点后面是否跟着字母或数字
            if (pos + 1 >= text.Length)
            {
                return false;
            }

            // 检查点后面是否有字母或数字（文件扩展名）
            int extensionStart = pos + 1;
            int extensionLength = 0;

            // 计算扩展名长度（最多5个字符，支持 .xlsm 这种4字符的）
            while (extensionStart + extensionLength < text.Length &&
                   extensionLength < 5 &&
                   char.IsLetterOrDigit(text[extensionStart + extensionLength]))
            {
                extensionLength++;
            }

            // 文件扩展名长度通常在2-5个字符之间
            if (extensionLength >= 2 && extensionLength <= 5)
            {
                // 检查扩展名后是否是空格、换行符、结束或其他非字母数字字符
                int nextPos = extensionStart + extensionLength;
                if (nextPos >= text.Length)
                {
                    return true; // 文本末尾
                }
                char charAfter = text[nextPos];
                // 如果后面是空格、换行符、标点符号或其他非字母数字字符，可能是文件后缀
                if (charAfter == ' ' || charAfter == '\t' || charAfter == '\r' || charAfter == '\n' ||
                    charAfter == ')' || charAfter == ']' || charAfter == '}' || charAfter == ',' ||
                    charAfter == ';' || charAfter == ':' || charAfter == '!' || charAfter == '?' ||
                    charAfter == '.' || charAfter == '"' || charAfter == '\'' || charAfter == '`')
                {
                    return true;
                }
                // 如果后面是中文或其他非ASCII字符，也可能是文件后缀的边界
                if (charAfter > 127)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 检查是否是缩写中的点（Mr. Mrs. Dr. i.e. e.g. U.S.等）
        /// </summary>
        private static bool IsAbbreviation(string text, int pos)
        {
            if (pos >= text.Length || text[pos] != '.')
            {
                return false;
            }

            // 常见缩写列表（不含点，用于向前匹配）
            string[] commonAbbrevs = { "Mr", "Mrs", "Ms", "Dr", "Prof", "etc", "Inc", "Ltd", "Jr", "Sr" };

            // 检查标准缩写（如 Mr.）
            foreach (string abbr in commonAbbrevs)
            {
                int start = pos - abbr.Length;
                if (start >= 0 && text.Substring(start, abbr.Length) == abbr)
                {
                    // 确保前面是单词边界（开头或非字母）
                    if (start == 0 || !char.IsLetter(text[start - 1]))
                    {
                        return true;
                    }
                }
            }

            // 检查 i.e 和 e.g 中的第一个点（i.e, e.g）
            if (pos > 0 && pos < text.Length - 1)
            {
                if (text[pos - 1] == 'i' && text[pos + 1] == 'e')
                {
                    return true; // i.e 中的点
                }
                if (text[pos - 1] == 'e' && text[pos + 1] == 'g')
                {
                    return true; // e.g 中的点
                }
            }

            // 检查 i.e. 和 e.g. 的最后一个点，以及 U.S. U.K. 等
            // 模式：X.Y.（前面是字母+点+字母）
            if (pos > 2)
            {
                if (char.IsLetter(text[pos - 1]) && text[pos - 2] == '.' && char.IsLetter(text[pos - 3]))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 检查字符串是否只包含空白字符（空格、\n、\r、\t等）
        /// </summary>
        /// <param name="s">待检查的字符串</param>
        /// <returns>如果只包含空白字符返回 true，否则返回 false</returns>
        private static bool IsWhitespaceOnly(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return true;
            }
            return string.IsNullOrWhiteSpace(s);
        }
    }
}

