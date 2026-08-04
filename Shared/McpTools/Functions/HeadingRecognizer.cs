using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    /// <summary>
    /// 按行分类结果：标题（按语言分组）与正文行（按换行切分后的非空非标题行）
    /// </summary>
    public sealed class DocumentLineClassificationResult
    {
        public DocumentLineClassificationResult(
            Dictionary<string, List<Dictionary<string, object>>> headingsByLanguage,
            List<Dictionary<string, object>> bodyLines)
        {
            HeadingsByLanguage = headingsByLanguage;
            BodyLines = bodyLines;
        }

        public Dictionary<string, List<Dictionary<string, object>>> HeadingsByLanguage { get; }
        public List<Dictionary<string, object>> BodyLines { get; }
    }

    /// <summary>
    /// 文本标题识别模块
    /// 通过正则表达式识别文本中的标题，并按类型分组
    /// 支持中英文标题模式和 Markdown 标题（#、##、### 等）
    /// </summary>
    public class HeadingRecognizer
    {
        // 中文标题模式
        private Dictionary<string, List<string>> chinesePatterns;
        
        // 英文标题模式
        private Dictionary<string, List<string>> englishPatterns;
        
        // 混合模式（中英文混合）
        private Dictionary<string, List<string>> mixedPatterns;
        
        // 标题关键词（用于辅助判断）
        private List<string> chineseKeywords;
        private List<string> englishKeywords;
        
        /// <summary>
        /// 初始化标题识别器，定义各种标题模式
        /// </summary>
        public HeadingRecognizer()
        {
            // 中文标题模式
            chinesePatterns = new Dictionary<string, List<string>>
            {
                {
                    "chapter", new List<string>
                    {
                        @"^\s*第[一二三四五六七八九十\d]+章[：:]?\s*",  // 第1章、第一章
                        @"^\s*第[一二三四五六七八九十\d]+节[：:]?\s*",  // 第1节
                    }
                },
                {
                    "numbered", new List<string>
                    {
                        @"^\s*[\d一二三四五六七八九十]+[\.、]\s*",  // 1.、1、一、
                        @"^\s*[\d一二三四五六七八九十]+\.\d+[\.、]?\s*",  // 1.1、1.1.
                        @"^\s*[\d一二三四五六七八九十]+\.\d+\.\d+[\.、]?\s*",  // 1.1.1
                        @"^\s*[（(][\d一二三四五六七八九十]+[)）]\s*",  // （1）、(一)
                        @"^\s*[①②③④⑤⑥⑦⑧⑨⑩]\s*",  // 圆圈数字
                    }
                },
                {
                    "keyword", new List<string>
                    {
                        @"^\s*(总结|结论|引言|背景|方法|结果|讨论|建议|概述|摘要|前言|目录|附录)",
                        @"^\s*(第一章|第二章|第三章|第四章|第五章|第六章|第七章|第八章|第九章|第十章)",
                    }
                },
            };
            
            // 英文标题模式
            englishPatterns = new Dictionary<string, List<string>>
            {
                {
                    "chapter", new List<string>
                    {
                        @"^\s*Chapter\s+[IVX\d]+[:\s]*",  // Chapter 1, Chapter I
                        @"^\s*CHAPTER\s+[IVX\d]+[:\s]*",  // CHAPTER 1
                        @"^\s*Part\s+[IVX\d]+[:\s]*",  // Part 1, Part I
                        @"^\s*PART\s+[IVX\d]+[:\s]*",  // PART 1
                        @"^\s*Section\s+[\d\.]+[:\s]*",  // Section 1, Section 1.1
                        @"^\s*SECTION\s+[\d\.]+[:\s]*",  // SECTION 1
                    }
                },
                {
                    "numbered", new List<string>
                    {
                        @"^\s*\d+[\.\)]\s+",  // 1. 1)
                        @"^\s*\d+\.\d+[\.\)]?\s+",  // 1.1, 1.1.
                        @"^\s*\d+\.\d+\.\d+[\.\)]?\s+",  // 1.1.1
                        @"^\s*[A-Z][\.\)]\s+",  // A. A)
                        @"^\s*[a-z][\.\)]\s+",  // a. a)
                        @"^\s*[IVX]+[\.\)]\s+",  // I. II. III.
                        @"^\s*[ivx]+[\.\)]\s+",  // i. ii. iii.
                        @"^\s*\([A-Z\d]+\)\s+",  // (A), (1)
                    }
                },
                {
                    "keyword", new List<string>
                    {
                        @"^\s*(Summary|Conclusion|Introduction|Background|Method|Methods|Results|Discussion|Recommendations|Abstract|Overview|Preface|Contents|Appendix)",
                        @"^\s*(SUMMARY|CONCLUSION|INTRODUCTION|BACKGROUND|METHOD|METHODS|RESULTS|DISCUSSION|RECOMMENDATIONS|ABSTRACT|OVERVIEW|PREFACE|CONTENTS|APPENDIX)",
                    }
                },
            };
            
            // 混合模式（中英文混合）
            mixedPatterns = new Dictionary<string, List<string>>
            {
                {
                    "numbered", new List<string>
                    {
                        @"^\s*[A-Z][\.、]\s*",  // A. A、
                        @"^\s*[a-z][\.、]\s*",  // a. a、
                    }
                },
            };
            
            // 标题关键词（用于辅助判断）
            chineseKeywords = new List<string>
            {
                "总结", "结论", "引言", "背景", "方法", "结果", "讨论", "建议",
                "概述", "摘要", "前言", "目录", "附录", "章节", "部分"
            };
            
            englishKeywords = new List<string>
            {
                "summary", "conclusion", "introduction", "background", "method", "methods",
                "results", "discussion", "recommendations", "abstract", "overview",
                "preface", "contents", "appendix", "chapter", "section", "part"
            };
        }
        
        /// <summary>
        /// 获取更详细的子类型分类
        /// </summary>
        /// <param name="text">标题文本</param>
        /// <param name="langType">语言类型 (chinese, english, mixed, markdown)</param>
        /// <param name="headingType">基础类型 (chapter, numbered, keyword, markdown)</param>
        /// <returns>详细的子类型</returns>
        public string GetDetailedSubtype(string text, string langType, string headingType)
        {
            string trimmedText = text.Trim();
            
            // Markdown 标题统一归类为 markdown_heading，不区分层级
            if (headingType == "markdown" || langType == "markdown")
            {
                return "markdown_heading";
            }
            
            if (headingType == "chapter")
            {
                return $"{langType}_chapter";
            }
            
            if (headingType == "keyword")
            {
                return $"{langType}_keyword";
            }
            
            if (headingType == "numbered")
            {
                // 中文序号类型的细分
                if (langType == "chinese")
                {
                    // 三级序号：1.1.1
                    if (Regex.IsMatch(trimmedText, @"^\d+\.\d+\.\d+"))
                    {
                        return "chinese_numbered_level3";
                    }
                    // 二级序号：1.1
                    else if (Regex.IsMatch(trimmedText, @"^\d+\.\d+"))
                    {
                        return "chinese_numbered_level2";
                    }
                    // 带括号的中文序号：（一）
                    else if (Regex.IsMatch(trimmedText, @"^[（(][一二三四五六七八九十]+[)）]"))
                    {
                        return "chinese_numbered_bracketed_chinese";
                    }
                    // 带括号的阿拉伯数字序号：（1）
                    else if (Regex.IsMatch(trimmedText, @"^[（(]\d+[)）]"))
                    {
                        return "chinese_numbered_bracketed_arabic";
                    }
                    // 中文数字序号：一、
                    else if (Regex.IsMatch(trimmedText, @"^[一二三四五六七八九十]+[、]"))
                    {
                        return "chinese_numbered_chinese_digit";
                    }
                    // 阿拉伯数字序号：1.
                    else if (Regex.IsMatch(trimmedText, @"^\d+[\.、]"))
                    {
                        return "chinese_numbered_arabic_digit";
                    }
                    // 圆圈数字：①②③
                    else if (Regex.IsMatch(trimmedText, @"^[①②③④⑤⑥⑦⑧⑨⑩]"))
                    {
                        return "chinese_numbered_circle";
                    }
                }
                
                // 英文序号类型的细分
                else if (langType == "english")
                {
                    // 三级序号：1.1.1
                    if (Regex.IsMatch(trimmedText, @"^\d+\.\d+\.\d+"))
                    {
                        return "english_numbered_level3";
                    }
                    // 二级序号：1.1
                    else if (Regex.IsMatch(trimmedText, @"^\d+\.\d+"))
                    {
                        return "english_numbered_level2";
                    }
                    // 罗马数字：I. II. III.
                    else if (Regex.IsMatch(trimmedText, @"^[IVX]+[\.\)]"))
                    {
                        return "english_numbered_roman";
                    }
                    // 大写字母：A. B.
                    else if (Regex.IsMatch(trimmedText, @"^[A-Z][\.\)]"))
                    {
                        return "english_numbered_uppercase";
                    }
                    // 小写字母：a. b.
                    else if (Regex.IsMatch(trimmedText, @"^[a-z][\.\)]"))
                    {
                        return "english_numbered_lowercase";
                    }
                    // 带括号：(A), (1)
                    else if (Regex.IsMatch(trimmedText, @"^\([A-Z\d]+\)"))
                    {
                        return "english_numbered_bracketed";
                    }
                    // 阿拉伯数字：1. 1)
                    else if (Regex.IsMatch(trimmedText, @"^\d+[\.\)]"))
                    {
                        return "english_numbered_arabic";
                    }
                }
                
                // 混合类型
                else if (langType == "mixed")
                {
                    return "mixed_numbered";
                }
            }
            
            return $"{langType}_{headingType}";
        }
        
        /// <summary>
        /// 判断文本是否为标题
        /// </summary>
        /// <param name="text">待判断的文本</param>
        /// <returns>元组：(是否为标题, 标题类型, 匹配的模式)</returns>
        public Tuple<bool, string, string> IsHeading(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return Tuple.Create(false, "", "");
            }

            // 仅去掉末尾空白，保留行首空格/制表符，便于 ^\s* 与序号之间的前导空白匹配
            string matchText = text.TrimEnd();
            string trimmedText = matchText.Trim();
            if (string.IsNullOrEmpty(trimmedText) || trimmedText.Length > 100)
            {
                return Tuple.Create(false, "", "");
            }

            // 先检查 Markdown 标题（#、##、### 等）
            string markdownPattern = @"^\s*#{1,6}\s+.+";
            if (Regex.IsMatch(matchText, markdownPattern))
            {
                return Tuple.Create(true, "markdown", "markdown");
            }
            
            // 先检查英文模式（优先级更高，因为数字模式可能重叠）
            foreach (var kvp in englishPatterns)
            {
                string headingType = kvp.Key;
                List<string> patterns = kvp.Value;
                
                foreach (string pattern in patterns)
                {
                    if (Regex.IsMatch(matchText, pattern, RegexOptions.IgnoreCase))
                    {
                        // 对于数字开头的模式，检查后续文本是否主要是英文
                        if (headingType == "numbered" && Regex.IsMatch(matchText, @"^\s*\d+"))
                        {
                            // 提取数字后的文本部分
                            string remaining = Regex.Replace(matchText, @"^\s*\d+[\.\)\s]+", "", RegexOptions.None);
                            // 如果包含中文字符，则可能是中文标题
                            if (Regex.IsMatch(remaining, @"[\u4e00-\u9fff]"))
                            {
                                continue;  // 跳过，让中文模式处理
                            }
                        }
                        return Tuple.Create(true, "english", headingType);
                    }
                }
            }
            
            // 检查中文模式
            foreach (var kvp in chinesePatterns)
            {
                string headingType = kvp.Key;
                List<string> patterns = kvp.Value;
                
                foreach (string pattern in patterns)
                {
                    if (Regex.IsMatch(matchText, pattern, RegexOptions.IgnoreCase))
                    {
                        return Tuple.Create(true, "chinese", headingType);
                    }
                }
            }
            
            // 检查混合模式
            foreach (var kvp in mixedPatterns)
            {
                string headingType = kvp.Key;
                List<string> patterns = kvp.Value;
                
                foreach (string pattern in patterns)
                {
                    if (Regex.IsMatch(matchText, pattern, RegexOptions.IgnoreCase))
                    {
                        return Tuple.Create(true, "mixed", headingType);
                    }
                }
            }
            
            // 检查关键词（作为辅助判断，需要更严格的条件）
            string textLower = matchText.ToLower();
            
            // 中文关键词检查：关键词必须在开头（允许前导空白），且文本较短
            foreach (string keyword in chineseKeywords)
            {
                if (Regex.IsMatch(matchText, @"^\s*" + Regex.Escape(keyword)) && trimmedText.Length < 30)
                {
                    // 排除包含"的"、"是"等常见句子结构的情况
                    if (!Regex.IsMatch(trimmedText, @"[的是了]\s*$"))
                    {
                        return Tuple.Create(true, "chinese", "keyword");
                    }
                }
            }
            
            // 英文关键词检查：关键词必须在开头，且是完整单词
            foreach (string keyword in englishKeywords)
            {
                // 检查关键词是否在文本开头（作为独立单词）
                string pattern = @"^\s*" + keyword + @"\b";
                if (Regex.IsMatch(textLower, pattern) && trimmedText.Length < 50)
                {
                    return Tuple.Create(true, "english", "keyword");
                }
            }
            
            return Tuple.Create(false, "", "");
        }
        
        /// <summary>
        /// 根据统计特征（数量和分布）判断每个标题的层级
        /// 
        /// 三步流程：
        /// 1. 识别标题，分出不同类别
        /// 2. 根据数量和分布情况加权赋分
        /// 3. 根据得分排序，每一类只能是一个层级
        /// </summary>
        /// <param name="allHeadings">所有标题列表，每个标题包含 'text', 'line', 'detailed_subtype' 等信息</param>
        /// <returns>标题文本到层级的映射字典</returns>
        public Dictionary<string, string> CalculateLevelByStatistics(List<Dictionary<string, object>> allHeadings)
        {
            if (allHeadings == null || allHeadings.Count == 0)
            {
                return new Dictionary<string, string>();
            }
            
            // ========== 第一步：识别标题，分出不同类别 ==========
            Dictionary<string, List<Dictionary<string, object>>> subtypeGroups = new Dictionary<string, List<Dictionary<string, object>>>();
            foreach (var heading in allHeadings)
            {
                string detailedSubtype = heading.ContainsKey("detailed_subtype") ? heading["detailed_subtype"]?.ToString() ?? "unknown" : "unknown";
                
                if (!subtypeGroups.ContainsKey(detailedSubtype))
                {
                    subtypeGroups[detailedSubtype] = new List<Dictionary<string, object>>();
                }
                subtypeGroups[detailedSubtype].Add(heading);
            }
            
            // ========== 第二步：根据数量和分布情况加权赋分 ==========
            Dictionary<string, double> subtypeScores = new Dictionary<string, double>();
            
            // 获取所有标题的行号范围（用于归一化）
            List<int> allLineNumbers = allHeadings
                .Where(h => h.ContainsKey("line") && h["line"] != null)
                .Select(h => Convert.ToInt32(h["line"]))
                .ToList();
            
            int totalRange = allLineNumbers.Count > 0 
                ? (allLineNumbers.Max() - allLineNumbers.Min()) 
                : 1;
            
            foreach (var kvp in subtypeGroups)
            {
                string subtype = kvp.Key;
                List<Dictionary<string, object>> headings = kvp.Value;
                
                if (headings == null || headings.Count == 0)
                {
                    continue;
                }
                
                List<int> lineNumbers = headings
                    .Where(h => h.ContainsKey("line") && h["line"] != null)
                    .Select(h => Convert.ToInt32(h["line"]))
                    .OrderBy(x => x)
                    .ToList();
                
                int count = headings.Count;
                int totalHeadings = allHeadings.Count;
                
                // 1. 数量得分：数量越少，得分越高
                double normalizedCountScore;
                if (count > 0 && totalHeadings > 0)
                {
                    double logCountScore = Math.Log((double)totalHeadings / count);
                    double maxLogScore = totalHeadings > 1 ? Math.Log((double)totalHeadings / 1) : 1;
                    double minLogScore = totalHeadings > 0 ? Math.Log((double)totalHeadings / totalHeadings) : 0;
                    
                    if (maxLogScore > minLogScore)
                    {
                        normalizedCountScore = (logCountScore - minLogScore) / (maxLogScore - minLogScore);
                    }
                    else
                    {
                        normalizedCountScore = 1.0;
                    }
                }
                else
                {
                    normalizedCountScore = 1.0;
                }
                
                // 2. 分布得分：分布越分散，得分越高
                double distributionScore;
                if (lineNumbers.Count == 1)
                {
                    distributionScore = 0.5;
                }
                else
                {
                    int lineRange = lineNumbers.Max() - lineNumbers.Min();
                    List<int> gaps = new List<int>();
                    for (int i = 0; i < lineNumbers.Count - 1; i++)
                    {
                        gaps.Add(lineNumbers[i + 1] - lineNumbers[i]);
                    }
                    
                    double avgGap = gaps.Count > 0 ? gaps.Average() : 0;
                    
                    double gapStd;
                    double gapCv;
                    if (avgGap > 0)
                    {
                        double variance = gaps.Count > 1 
                            ? gaps.Select(x => Math.Pow(x - avgGap, 2)).Sum() / (gaps.Count - 1)
                            : 0;
                        gapStd = Math.Sqrt(variance);
                        gapCv = gapStd / avgGap;
                    }
                    else
                    {
                        gapCv = 0;
                    }
                    
                    if (totalRange > 0)
                    {
                        double rangeScore = Math.Min((double)lineRange / totalRange, 1.0);
                        double avgGapScore = totalRange > 0 ? Math.Min(avgGap / totalRange, 1.0) : 0;
                        double cvScore = Math.Min(gapCv, 1.0);
                        distributionScore = 0.4 * rangeScore + 0.4 * avgGapScore + 0.2 * cvScore;
                    }
                    else
                    {
                        distributionScore = 0;
                    }
                }
                
                // 加权综合得分（数量和分布各占50%）
                double finalScore = 0.5 * normalizedCountScore + 0.5 * distributionScore;
                subtypeScores[subtype] = finalScore;
            }
            
            // ========== 第三步：根据得分排序，每一类只能是一个层级 ==========
            var sortedSubtypes = subtypeScores.OrderByDescending(x => x.Value).ToList();
            
            // 按得分排序，直接分配层级：第1个=Heading 1，第2个=Heading 2，第3个=Heading 3，...
            Dictionary<string, int> subtypeLevels = new Dictionary<string, int>();
            for (int idx = 0; idx < sortedSubtypes.Count; idx++)
            {
                string subtype = sortedSubtypes[idx].Key;
                int level = idx + 1;  // 按顺序分配1,2,3,4,5...
                subtypeLevels[subtype] = level;
            }
            
            // 为每个标题分配层级（同一类别的所有标题使用相同的层级）
            Dictionary<string, string> levelMapping = new Dictionary<string, string>();
            foreach (var heading in allHeadings)
            {
                string subtype = heading.ContainsKey("detailed_subtype") ? heading["detailed_subtype"]?.ToString() ?? "unknown" : "unknown";
                int level = subtypeLevels.ContainsKey(subtype) ? subtypeLevels[subtype] : 2;  // 默认层级2
                string text = heading.ContainsKey("text") ? heading["text"]?.ToString() ?? "" : "";
                levelMapping[text] = $"Heading {level}";
            }
            
            return levelMapping;
        }
        
        /// <summary>
        /// 根据特征判断标题级别
        /// 
        /// 判断规则（按优先级从高到低）：
        /// 1. 三级标题：多级序号（1.1.1、1.1.1.1）
        /// 2. 二级标题：
        ///    - 二级序号（1.1、1.2）
        ///    - 带括号序号（（一）、（1））
        ///    - Section 标题
        /// 3. 一级标题：
        ///    - 章节标题（第一章、Chapter 1、Part 1）
        ///    - 中文数字序号（一、二、三）
        ///    - 单独的阿拉伯数字序号（1. 2. 3.）- 如果前面没有中文数字序号
        /// 
        /// Heading 级别说明：
        /// - Heading 1: 一级标题（最顶层），如：第一章、一、项目背景、Chapter 1
        /// - Heading 2: 二级标题（一级标题下的子标题），如：（一）技术需求、1.1 研究目的
        /// - Heading 3: 三级标题（二级标题下的子标题），如：1.1.1 具体目标
        /// </summary>
        /// <param name="text">标题文本</param>
        /// <param name="headingType">标题类型</param>
        /// <param name="previousHeadings">之前的标题列表（用于上下文判断，可选）</param>
        /// <returns>标题级别 (Heading 1, Heading 2, Heading 3)</returns>
        public string AssignHeadingLevel(string text, string headingType, List<Dictionary<string, object>> previousHeadings = null)
        {
            string trimmedText = text.Trim();
            
            // ========== 第一优先级：三级标题 ==========
            // 多级序号：1.1.1、1.1.1.1 等
            if (Regex.IsMatch(trimmedText, @"^\d+\.\d+\.\d+"))
            {
                return "Heading 3";
            }
            
            // ========== 第二优先级：二级标题 ==========
            // 1. 二级序号：1.1、1.2、2.1 等
            if (Regex.IsMatch(trimmedText, @"^\d+\.\d+"))
            {
                return "Heading 2";
            }
            
            // 2. 带括号的中文序号：（一）、（二）、（三）
            if (Regex.IsMatch(trimmedText, @"^[（(][一二三四五六七八九十]+[)）]"))
            {
                return "Heading 2";
            }
            
            // 3. 带括号的阿拉伯数字序号：（1）、（2）、（3）
            if (Regex.IsMatch(trimmedText, @"^[（(]\d+[)）]"))
            {
                return "Heading 2";
            }
            
            // 4. Section 标题：Section 1.1、Section 2.1
            if (Regex.IsMatch(trimmedText, @"^Section\s+", RegexOptions.IgnoreCase))
            {
                return "Heading 2";
            }
            
            // ========== 第三优先级：一级标题 ==========
            // 1. 章节标题：第一章、Chapter 1、Part 1
            if (Regex.IsMatch(trimmedText, @".*[章Chapter|PART|Part].*", RegexOptions.IgnoreCase))
            {
                return "Heading 1";
            }
            
            // 2. 中文数字序号：一、二、三、四（顶层序号）
            if (Regex.IsMatch(trimmedText, @"^[一二三四五六七八九十]+[、]"))
            {
                return "Heading 1";
            }
            
            // 3. Chapter、Part 开头的英文标题
            if (Regex.IsMatch(trimmedText, @"^Chapter\s+", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(trimmedText, @"^Part\s+", RegexOptions.IgnoreCase))
            {
                return "Heading 1";
            }
            
            // ========== 第四优先级：根据上下文判断 ==========
            // 对于单独的阿拉伯数字序号（1. 2. 3.），需要根据上下文判断
            if (Regex.IsMatch(trimmedText, @"^\d+[\.、]"))
            {
                // 如果提供了之前的标题列表，检查上下文
                if (previousHeadings != null && previousHeadings.Count > 0)
                {
                    // 检查最近的几个标题，看是否有中文数字序号（一、二、三）
                    int checkCount = Math.Min(5, previousHeadings.Count);
                    for (int i = previousHeadings.Count - checkCount; i < previousHeadings.Count; i++)
                    {
                        var prevHeading = previousHeadings[i];
                        string prevText = prevHeading.ContainsKey("text") ? prevHeading["text"]?.ToString() ?? "" : "";
                        
                        // 如果前面有中文数字序号，则当前阿拉伯数字序号可能是二级标题
                        if (Regex.IsMatch(prevText, @"^[一二三四五六七八九十]+[、]"))
                        {
                            return "Heading 2";
                        }
                        // 如果前面有带括号的序号，则当前可能是三级标题
                        if (Regex.IsMatch(prevText, @"^[（(]"))
                        {
                            return "Heading 3";
                        }
                    }
                }
                // 默认作为一级标题
                return "Heading 1";
            }
            
            // ========== 默认：一级标题 ==========
            return "Heading 1";
        }
        
        
        
        /// <summary>
        /// 迭代删除「叶子表」：成对 <c>&lt;table&gt;…&lt;/table&gt;</c> 内部不再含嵌套 <c>&lt;table</c> 的整块。
        /// 自内向外反复删除，直至无叶子表（嵌套表会先去掉内层，外层再变为叶子）。
        /// </summary>
        private static string RemoveLeafTableBlocks(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            // 最内层表先被匹配：从第一个 <table> 起，内容不能出现嵌套 <table（(?!<table) 逐字约束）
            const string pattern = @"<table(?:\s+[^>]*)?>((?:(?!<table)[\s\S])*?)</table\s*>";
            var regex = new Regex(pattern, RegexOptions.IgnoreCase);
            string s = text;
            while (regex.IsMatch(s))
            {
                s = regex.Replace(s, "", 1);
            }

            return s;
        }

        /// <summary>
        /// 切行扫描前：先移除所有叶子 <c>&lt;table&gt;…&lt;/table&gt;</c>（自内向外删尽），再在每个 <c>&lt;/cell&gt;</c> 前插入 <c>\r</c>，再移除 <c>&lt;image /&gt;</c>、表格结构标签，便于按行识别标题。
        /// </summary>
        private static string PrepareTextForHeadingLineScan(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string s = RemoveLeafTableBlocks(text);

            // 先在每个 </cell> 前补 \r，再删标签（保证单元格边界在分割时成为独立行）
            s = Regex.Replace(s, @"</cell\s*>", "\r</cell>", RegexOptions.IgnoreCase);
            // <image .../>、<table ...>、<row>、<cell ...>、及对应闭合标签
            s = Regex.Replace(
                s,
                @"<image(?:\s+[^>]*)?/>|<table(?:\s+[^>]*)?>|<row(?:\s+[^>]*)?>|<cell(?:\s+[^>]*)?>|</table\s*>|</row\s*>|</cell\s*>",
                "",
                RegexOptions.IgnoreCase);
            return s;
        }

        /// <summary>
        /// 按 \r\n、\r、\n 切行后扫描：标题行加入 allHeadings，其余非空行加入 bodyLines。
        /// 切行前会在每个 <c>&lt;/cell&gt;</c> 前插入 <c>\r</c>，再移除表格/图片相关标签，避免标签与正文黏在同一行影响标题识别。
        /// 标题条目的 <c>text</c> 存切分后的原始行（与 python heading_recognizer 一致）；判题与子类型仍用去首尾空白后的串。
        /// </summary>
        private void ScanLinesForHeadingsAndBody(string text, out List<Dictionary<string, object>> allHeadings, out List<Dictionary<string, object>> bodyLines)
        {
            text = PrepareTextForHeadingLineScan(text);
            string[] lines = Regex.Split(text, @"\r\n|\r|\n");
            allHeadings = new List<Dictionary<string, object>>();
            bodyLines = new List<Dictionary<string, object>>();

            Debug.WriteLine("");
            Debug.WriteLine("========== [HeadingRecognizer] 经去标签与 </cell> 前补 \\r 后，按 Regex.Split(@\"\\r\\n|\\r|\\n\") 分割 ==========");
            Debug.WriteLine($"分割段总数: {lines.Length}（含空段；\\r\\n 与 \\r、\\n 均为分隔符）");

            for (int lineNum = 0; lineNum < lines.Length; lineNum++)
            {
                string line = lines[lineNum];
                string trimmed = line.Trim();
                string escaped = TextUtils.EscapeNewlinesForDb(line);

                if (string.IsNullOrEmpty(trimmed))
                {
                    Debug.WriteLine($"  seg#{lineNum + 1,4}  len={line.Length,-5}  (空段，跳过)  |  {escaped}");
                    continue;
                }

                // 须保留行首空白再判标题：如「     （二）融入思路…」依赖 ^\s*；勿传 trimmed（会去掉行首缩进）。
                // 句中「似乎没有把（二）…」不要求识别，整行 ^ 不匹配括号序号即可。
                var isHeadingResult = IsHeading(line.TrimEnd());
                bool isHeadingFlag = isHeadingResult.Item1;
                string langType = isHeadingResult.Item2;
                string headingType = isHeadingResult.Item3;

                if (isHeadingFlag)
                {
                    string detailedSubtype = GetDetailedSubtype(trimmed, langType, headingType);
                    allHeadings.Add(new Dictionary<string, object>
                    {
                        { "text", line },
                        { "line", lineNum + 1 },
                        { "language", langType },
                        { "subtype", headingType },
                        { "detailed_subtype", detailedSubtype }
                    });
                    Debug.WriteLine($"  seg#{lineNum + 1,4}  len={line.Length,-5}  【标题】 {langType}/{headingType}  detailed={detailedSubtype}  |  {escaped}");
                }
                else
                {
                    bodyLines.Add(new Dictionary<string, object>
                    {
                        { "type", "body" },
                        { "text", trimmed },
                        { "line", lineNum + 1 }
                    });
                    Debug.WriteLine($"  seg#{lineNum + 1,4}  len={line.Length,-5}  【正文】  |  {escaped}");
                }
            }

            Debug.WriteLine($"========== [HeadingRecognizer] 分割结束：标题行 {allHeadings.Count}，正文行 {bodyLines.Count} ==========");
            Debug.WriteLine("");
        }

        private Dictionary<string, List<Dictionary<string, object>>> ApplyLevelAndGroupHeadingsByLanguage(List<Dictionary<string, object>> allHeadings)
        {
            Dictionary<string, string> levelMapping = CalculateLevelByStatistics(allHeadings);
            Dictionary<string, List<Dictionary<string, object>>> headingsByType = new Dictionary<string, List<Dictionary<string, object>>>();
            foreach (var heading in allHeadings)
            {
                string headingText = heading.ContainsKey("text") ? heading["text"]?.ToString() ?? "" : "";
                string level = levelMapping.ContainsKey(headingText) ? levelMapping[headingText] : "Heading 1";
                heading["level"] = level;

                string language = heading.ContainsKey("language") ? heading["language"]?.ToString() ?? "" : "";
                if (!headingsByType.ContainsKey(language))
                {
                    headingsByType[language] = new List<Dictionary<string, object>>();
                }
                headingsByType[language].Add(heading);
            }

            return headingsByType;
        }

        /// <summary>
        /// 对全文按行分类：标题（按语言分组并带 level）与正文行，供格式提取管线使用
        /// </summary>
        public DocumentLineClassificationResult ClassifyLinesForFormatExtraction(string text)
        {
            ScanLinesForHeadingsAndBody(text, out List<Dictionary<string, object>> allHeadings, out List<Dictionary<string, object>> bodyLines);
            return new DocumentLineClassificationResult(ApplyLevelAndGroupHeadingsByLanguage(allHeadings), bodyLines);
        }

        public Dictionary<string, List<Dictionary<string, object>>> RecognizeHeadings(string text)
        {
            ScanLinesForHeadingsAndBody(text, out List<Dictionary<string, object>> allHeadings, out List<Dictionary<string, object>> _);
            return ApplyLevelAndGroupHeadingsByLanguage(allHeadings);
        }
        
        
        
        /// <summary>
        /// 获取标题级别的中文说明
        /// </summary>
        /// <param name="level">标题级别 (Heading 1, Heading 2, Heading 3)</param>
        /// <returns>中文说明</returns>
        public string GetLevelDescription(string level)
        {
            Dictionary<string, string> descriptions = new Dictionary<string, string>
            {
                { "Heading 1", "一级标题（最顶层）" },
                { "Heading 2", "二级标题（一级标题下的子标题）" },
                { "Heading 3", "三级标题（二级标题下的子标题）" },
            };
            
            return descriptions.ContainsKey(level) ? descriptions[level] : level;
        }
        
        /// <summary>
        /// 获取子类型的中文说明
        /// </summary>
        /// <param name="subtype">子类型名称</param>
        /// <returns>中文说明</returns>
        public string GetSubtypeDescription(string subtype)
        {
            Dictionary<string, string> descriptions = new Dictionary<string, string>
            {
                { "chinese_chapter", "中文章节标题（第X章）" },
                { "chinese_numbered_chinese_digit", "中文数字序号（一、二、三）" },
                { "chinese_numbered_arabic_digit", "阿拉伯数字序号（1. 2. 3.）" },
                { "chinese_numbered_bracketed_chinese", "带括号中文序号（（一）（二））" },
                { "chinese_numbered_bracketed_arabic", "带括号阿拉伯数字序号（（1）（2））" },
                { "chinese_numbered_level2", "二级序号（1.1 1.2）" },
                { "chinese_numbered_level3", "三级序号（1.1.1）" },
                { "chinese_numbered_circle", "圆圈数字序号（①②③）" },
                { "chinese_keyword", "中文关键词标题" },
                { "english_chapter", "英文章节标题（Chapter X）" },
                { "english_numbered_arabic", "英文阿拉伯数字序号（1. 1)）" },
                { "english_numbered_level2", "英文二级序号（1.1）" },
                { "english_numbered_level3", "英文三级序号（1.1.1）" },
                { "english_numbered_roman", "罗马数字序号（I. II. III.）" },
                { "english_numbered_uppercase", "大写字母序号（A. B.）" },
                { "english_numbered_lowercase", "小写字母序号（a. b.）" },
                { "english_numbered_bracketed", "带括号英文序号（(A) (1)）" },
                { "english_keyword", "英文关键词标题" },
                { "mixed_numbered", "混合序号" },
                { "markdown_heading", "Markdown标题（# ## ###）" },
            };
            
            return descriptions.ContainsKey(subtype) ? descriptions[subtype] : subtype;
        }
        
        /// <summary>
        /// 打印识别到的标题
        /// </summary>
        /// <param name="headings">标题字典</param>
        /// <param name="groupBySubtype">是否按子类型分组显示</param>
        public void PrintHeadings(Dictionary<string, List<Dictionary<string, object>>> headings, bool groupBySubtype = false)
        {
            if (headings == null || headings.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("未识别到任何标题");
                return;
            }
            
            if (groupBySubtype)
            {
                // 按子类型排序，中文在前，英文在后
                var sortedSubtypes = headings.Keys.OrderBy(x => 
                    x.StartsWith("chinese") ? 0 : (x.StartsWith("english") ? 1 : 2)
                ).ThenBy(x => x).ToList();
                
                foreach (string subtype in sortedSubtypes)
                {
                    var headingList = headings[subtype];
                    string description = GetSubtypeDescription(subtype);
                    System.Diagnostics.Debug.WriteLine($"\n【{description}】({headingList.Count}个)");
                    System.Diagnostics.Debug.WriteLine(new string('-', 60));
                    
                    foreach (var heading in headingList)
                    {
                        int line = heading.ContainsKey("line") ? Convert.ToInt32(heading["line"]) : 0;
                        string level = heading.ContainsKey("level") ? heading["level"]?.ToString() ?? "" : "";
                        string levelDesc = GetLevelDescription(level);
                        string text = heading.ContainsKey("text") ? heading["text"]?.ToString() ?? "" : "";
                        System.Diagnostics.Debug.WriteLine($"  行{line,4} [{level,-12} ({levelDesc})] {text}");
                    }
                }
            }
            else
            {
                foreach (var kvp in headings)
                {
                    string langType = kvp.Key;
                    var headingList = kvp.Value;
                    System.Diagnostics.Debug.WriteLine($"\n【{langType.ToUpper()}标题】({headingList.Count}个)");
                    System.Diagnostics.Debug.WriteLine(new string('-', 60));
                    
                    foreach (var heading in headingList)
                    {
                        int line = heading.ContainsKey("line") ? Convert.ToInt32(heading["line"]) : 0;
                        string level = heading.ContainsKey("level") ? heading["level"]?.ToString() ?? "" : "";
                        string levelDesc = GetLevelDescription(level);
                        string text = heading.ContainsKey("text") ? heading["text"]?.ToString() ?? "" : "";
                        System.Diagnostics.Debug.WriteLine($"  行{line,4} [{level,-12} ({levelDesc})] {text}");
                    }
                }
            }
        }
    }
}

