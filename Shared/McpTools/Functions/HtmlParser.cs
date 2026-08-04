using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// HTML解析器，用于解析HTML内容并设置Word格式
    /// </summary>
    public static class HtmlParser
    {
        /// <summary>
        /// HTML文本段（包含文本内容和格式信息）
        /// </summary>
        public class HtmlTextSegment
        {
            public string Text { get; set; }
            public string Color { get; set; }
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
        }

        /// <summary>
        /// 解析HTML内容，提取文本段和格式信息
        /// </summary>
        /// <param name="htmlContent">HTML内容（如：&lt;span style='color:red;font-weight:bold'&gt;姓名&lt;/span&gt;）</param>
        /// <returns>文本段列表</returns>
        public static List<HtmlTextSegment> ParseHtml(string htmlContent)
        {
            var segments = new List<HtmlTextSegment>();
            
            if (string.IsNullOrEmpty(htmlContent))
            {
                return segments;
            }

            try
            {
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] 🔍 开始解析HTML内容: '{htmlContent}'");
                
                // 使用正则表达式匹配 <span> 标签
                // 匹配格式：<span style='...'>文本</span>
                string pattern = @"<span\s+style=['""]([^'""]*)['""]\s*>([^<]*)</span>";
                var matches = Regex.Matches(htmlContent, pattern, RegexOptions.IgnoreCase);
                
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] 📊 正则匹配结果: 找到 {matches.Count} 个span标签");

                int lastIndex = 0;
                foreach (Match match in matches)
                {
                    // 处理匹配前的文本（如果有）
                    if (match.Index > lastIndex)
                    {
                        string beforeText = htmlContent.Substring(lastIndex, match.Index - lastIndex);
                        // 移除HTML标签
                        beforeText = Regex.Replace(beforeText, "<[^>]+>", "");
                        if (!string.IsNullOrEmpty(beforeText))
                        {
                            segments.Add(new HtmlTextSegment
                            {
                                Text = beforeText,
                                Color = null,
                                IsBold = false,
                                IsItalic = false
                            });
                        }
                    }

                    // 解析style属性
                    string style = match.Groups[1].Value;
                    string text = match.Groups[2].Value;
                    
                    bool isBold = ParseBoldFromStyle(style);
                    bool isItalic = ParseItalicFromStyle(style);
                    string color = ParseColorFromStyle(style);
                    
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser] ✅ 解析到span标签:");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - style属性: '{style}'");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - 文本内容: '{text}'");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - IsBold: {isBold}");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - IsItalic: {isItalic}");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - Color: {color ?? "null"}");
                    
                    var segment = new HtmlTextSegment
                    {
                        Text = text,
                        Color = color,
                        IsBold = isBold,
                        IsItalic = isItalic
                    };
                    
                    segments.Add(segment);
                    lastIndex = match.Index + match.Length;
                }

                // 处理剩余的文本（如果有）
                if (lastIndex < htmlContent.Length)
                {
                    string remainingText = htmlContent.Substring(lastIndex);
                    // 移除HTML标签
                    remainingText = Regex.Replace(remainingText, "<[^>]+>", "");
                    if (!string.IsNullOrEmpty(remainingText))
                    {
                        segments.Add(new HtmlTextSegment
                        {
                            Text = remainingText,
                            Color = null,
                            IsBold = false,
                            IsItalic = false
                        });
                    }
                }

                // 如果没有匹配到任何span标签，直接返回纯文本
                if (segments.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser] ⚠️ 未匹配到任何span标签，降级为纯文本");
                    string plainText = Regex.Replace(htmlContent, "<[^>]+>", "");
                    if (!string.IsNullOrEmpty(plainText))
                    {
                        segments.Add(new HtmlTextSegment
                        {
                            Text = plainText,
                            Color = null,
                            IsBold = false,
                            IsItalic = false
                        });
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] 📋 解析完成，共生成 {segments.Count} 个文本段");
                for (int i = 0; i < segments.Count; i++)
                {
                    var seg = segments[i];
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   段{i + 1}: 文本='{seg.Text}', IsBold={seg.IsBold}, IsItalic={seg.IsItalic}, Color={seg.Color ?? "null"}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] ❌ 解析HTML失败: {ex.Message}");
                // 降级为纯文本
                string plainText = Regex.Replace(htmlContent, "<[^>]+>", "");
                if (!string.IsNullOrEmpty(plainText))
                {
                    segments.Add(new HtmlTextSegment
                    {
                        Text = plainText,
                        Color = null,
                        IsBold = false,
                        IsItalic = false
                    });
                }
            }

            return segments;
        }

        /// <summary>
        /// 从style属性中解析颜色
        /// </summary>
        private static string ParseColorFromStyle(string style)
        {
            if (string.IsNullOrEmpty(style))
                return null;

            // 匹配 color:value 或 color: value
            var match = Regex.Match(style, @"color\s*:\s*([^;]+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string colorValue = match.Groups[1].Value.Trim();
                
                // 处理颜色名称（如 red, blue 等）
                if (IsColorName(colorValue))
                {
                    return ColorNameToHex(colorValue);
                }
                
                // 处理十六进制颜色（如 #FF0000）
                if (colorValue.StartsWith("#"))
                {
                    return colorValue.Substring(1); // 移除 #
                }
                
                // 处理rgb格式（如 rgb(255,0,0)）
                var rgbMatch = Regex.Match(colorValue, @"rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", RegexOptions.IgnoreCase);
                if (rgbMatch.Success)
                {
                    int r = int.Parse(rgbMatch.Groups[1].Value);
                    int g = int.Parse(rgbMatch.Groups[2].Value);
                    int b = int.Parse(rgbMatch.Groups[3].Value);
                    return $"{r:X2}{g:X2}{b:X2}";
                }
                
                return colorValue;
            }

            return null;
        }

        /// <summary>
        /// 从style属性中解析粗体
        /// </summary>
        private static bool ParseBoldFromStyle(string style)
        {
            if (string.IsNullOrEmpty(style))
            {
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] ParseBoldFromStyle: style为空，返回false");
                return false;
            }

            // 匹配 font-weight:bold 或 font-weight: bold
            var match = Regex.Match(style, @"font-weight\s*:\s*(\w+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string weight = match.Groups[1].Value.Trim().ToLower();
                bool isBold = weight == "bold" || weight == "700" || weight == "800" || weight == "900";
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] ParseBoldFromStyle: 匹配到font-weight='{weight}', 结果={isBold}");
                return isBold;
            }

            System.Diagnostics.Debug.WriteLine($"[HtmlParser] ParseBoldFromStyle: 未匹配到font-weight，返回false");
            return false;
        }

        /// <summary>
        /// 从style属性中解析斜体
        /// </summary>
        private static bool ParseItalicFromStyle(string style)
        {
            if (string.IsNullOrEmpty(style))
                return false;

            // 匹配 font-style:italic 或 font-style: italic
            var match = Regex.Match(style, @"font-style\s*:\s*(\w+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string styleValue = match.Groups[1].Value.Trim().ToLower();
                return styleValue == "italic" || styleValue == "oblique";
            }

            return false;
        }

        /// <summary>
        /// 检查是否为颜色名称
        /// </summary>
        private static bool IsColorName(string color)
        {
            if (string.IsNullOrEmpty(color))
                return false;

            string[] colorNames = {
                "red", "green", "blue", "yellow", "orange", "purple", "pink",
                "black", "white", "gray", "grey", "cyan", "magenta", "brown",
                "navy", "maroon", "olive", "lime", "aqua", "silver", "gold"
            };

            return colorNames.Contains(color.ToLower());
        }

        /// <summary>
        /// 将颜色名称转换为十六进制
        /// </summary>
        private static string ColorNameToHex(string colorName)
        {
            var colorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "red", "FF0000" },
                { "green", "008000" },
                { "blue", "0000FF" },
                { "yellow", "FFFF00" },
                { "orange", "FFA500" },
                { "purple", "800080" },
                { "pink", "FFC0CB" },
                { "black", "000000" },
                { "white", "FFFFFF" },
                { "gray", "808080" },
                { "grey", "808080" },
                { "cyan", "00FFFF" },
                { "magenta", "FF00FF" },
                { "brown", "A52A2A" },
                { "navy", "000080" },
                { "maroon", "800000" },
                { "olive", "808000" },
                { "lime", "00FF00" },
                { "aqua", "00FFFF" },
                { "silver", "C0C0C0" },
                { "gold", "FFD700" }
            };

            return colorMap.TryGetValue(colorName, out string hex) ? hex : null;
        }

        /// <summary>
        /// 将RGB十六进制字符串转换为Word.WdColor
        /// </summary>
        public static Word.WdColor ConvertRgbToWdColor(string rgbString)
        {
            try
            {
                if (string.IsNullOrEmpty(rgbString) || rgbString == "automatic")
                {
                    return Word.WdColor.wdColorAutomatic;
                }

                // 移除 # 符号（如果有）
                string hex = rgbString.TrimStart('#');

                if (hex.Length == 6)
                {
                    // 解析RGB值
                    int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                    int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                    int b = Convert.ToInt32(hex.Substring(4, 2), 16);

                    // Word颜色值是BGR格式，需要转换
                    int bgrValue = (b << 16) | (g << 8) | r;

                    return (Word.WdColor)bgrValue;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] ❌ 转换颜色失败: {ex.Message}");
            }

            return Word.WdColor.wdColorAutomatic;
        }

        /// <summary>
        /// 将解析的HTML文本段插入到Word Range并设置格式
        /// </summary>
        /// <param name="range">Word Range对象</param>
        /// <param name="segments">HTML文本段列表</param>
        /// <returns>是否成功</returns>
        public static bool InsertHtmlSegments(Word.Range range, List<HtmlTextSegment> segments)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] 🚀 InsertHtmlSegments: 开始插入 {segments?.Count ?? 0} 个文本段");
                
                if (range == null || segments == null || segments.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser] ⚠️ InsertHtmlSegments: 参数无效，返回false");
                    return false;
                }

                // 记录插入前的起始位置
                int startPos = range.Start;
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] 📍 插入前Range位置: Start={startPos}, End={range.End}");
                
                // 先插入所有文本
                string allText = string.Join("", segments.Select(s => s.Text));
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] 📝 准备插入的完整文本: '{allText}'");
                range.Text = allText;

                // 设置Text后，range会扩展到新文本的范围，需要重新获取起始位置
                int newStartPos = range.Start;
                int currentPos = newStartPos;
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] 📍 插入后Range位置: Start={newStartPos}, End={range.End}");

                // 为每个文本段设置格式
                int segmentIndex = 0;
                foreach (var segment in segments)
                {
                    segmentIndex++;
                    if (string.IsNullOrEmpty(segment.Text))
                    {
                        System.Diagnostics.Debug.WriteLine($"[HtmlParser] ⏭️ 段{segmentIndex}: 文本为空，跳过");
                        continue;
                    }

                    int textLength = segment.Text.Length;
                    int segmentStart = currentPos;
                    int segmentEnd = currentPos + textLength;

                    System.Diagnostics.Debug.WriteLine($"[HtmlParser] 🎨 段{segmentIndex}: 设置格式");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - 文本: '{segment.Text}'");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - 位置: {segmentStart}-{segmentEnd}");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - IsBold: {segment.IsBold}");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - IsItalic: {segment.IsItalic}");
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   - Color: {segment.Color ?? "null"}");

                    // 创建文本段的Range
                    Word.Range segmentRange = range.Document.Range(segmentStart, segmentEnd);

                    // 设置颜色
                    if (!string.IsNullOrEmpty(segment.Color))
                    {
                        Word.WdColor wdColor = ConvertRgbToWdColor(segment.Color);
                        segmentRange.Font.Color = wdColor;
                        System.Diagnostics.Debug.WriteLine($"[HtmlParser]   ✅ 已设置颜色: {segment.Color}");
                    }

                    // 设置粗体
                    int boldValue = segment.IsBold ? 1 : 0;
                    segmentRange.Font.Bold = boldValue;
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   ✅ 已设置粗体: Font.Bold = {boldValue} (IsBold={segment.IsBold})");
                    
                    // 验证设置是否成功
                    int actualBold = segmentRange.Font.Bold;
                    System.Diagnostics.Debug.WriteLine($"[HtmlParser]   🔍 验证: 实际Font.Bold值 = {actualBold}");

                    // 设置斜体
                    segmentRange.Font.Italic = segment.IsItalic ? 1 : 0;

                    currentPos = segmentEnd;
                }

                System.Diagnostics.Debug.WriteLine($"[HtmlParser] ✅ InsertHtmlSegments: 完成，共处理 {segmentIndex} 个文本段");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HtmlParser] ❌ 插入HTML文本段失败: {ex.Message}");
                return false;
            }
        }
    }
}

