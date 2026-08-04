using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 获取Word文档内容工具
    /// 获取当前编辑的Word文档的完整内容
    /// </summary>
    public static class F_GetDocumentContentTool
    {
        /// <summary>
        /// 注册获取文档内容工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            List<McpTool> availableTools,
            object wordApplication)
        {
            var tool = new McpTool
            {
                name = "F_get_document_content",
                alias = "获取文档内容",
                description = "获取当前编辑的Word文档的完整内容",
                usage = "当需要获取当前Word文档内容时，使用此工具",
                input_schema = new McpTool.ToolInputSchema
                {
                    Type = "object",
                    Properties = new Dictionary<string, object>
                    {
                        ["format"] = new
                        {
                            type = "string",
                            description = "返回内容的格式",
                            @enum = new[] { "text", "html" },
                            @default = "text"
                        }
                    },
                    required = new List<string>() // 没有必需参数，format有默认值
                }
            };

            toolRegistry["F_get_document_content"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序实例不可用" };
                    }

                    // 转换为Word应用程序类型
                    dynamic wordApp = wordApplication;

                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    string format = args.ContainsKey("format") ? args["format"]?.ToString() : "text";

                    string content = "";
                    try
                    {
                        if (format == "html")
                        {
                            // 以HTML格式获取内容 - 直接使用Text属性
                            Word.Range contentRange = wordApp.ActiveDocument.Content;
                            content = contentRange.Text ?? "";

                            // 在HTML内容中插入表格和图片占位符
                            content = InsertElementPlaceholders(wordApp.ActiveDocument, content, "html");

                            // 简单的HTML包装
                            content = $"<html><body><pre>{System.Web.HttpUtility.HtmlEncode(content)}</pre></body></html>";
                        }
                        else
                        {
                            // 以纯文本格式获取内容 - 直接使用Text属性
                            Word.Range contentRange = wordApp.ActiveDocument.Content;
                            content = contentRange.Text ?? "";

                            // 在文本内容中插入表格和图片占位符
                            content = InsertElementPlaceholders(wordApp.ActiveDocument, content, "text");

                            // 过滤临时内容（删除线+特殊背景色，特殊黄色背景，按钮）
                            string filteredContent = FilterTempContent(wordApp.ActiveDocument, content, format);
                            content = filteredContent;
                        }
                    }
                    catch (Exception apiEx)
                    {
                        // 备用方法：尝试直接转换为字符串
                        try
                        {
                            content = wordApp.ActiveDocument.Content.ToString() ?? "";
                        }
                        catch (Exception)
                        {
                            throw new Exception($"无法获取文档内容: {apiEx.Message}");
                        }
                    }

                    await Task.CompletedTask; // 确保异步执行

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            content = content,
                            format = format,
                            length = content.Length,
                            document_name = wordApp.ActiveDocument.Name ?? "未命名文档"
                        }
                    };

                    return result;
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"获取文档内容失败: {ex.Message}" };
                }
            };

            availableTools.Add(tool);
        }

        /// <summary>
        /// 在文档内容中插入表格和图片占位符
        /// </summary>
        private static string InsertElementPlaceholders(Word.Document doc, string originalContent, string format)
        {
            try
            {
                // 收集所有需要插入的占位符信息
                var placeholders = new List<(int position, string placeholder, int length)>();

                // 收集表格占位符信息
                int tableIndex = 1;
                foreach (Word.Table table in doc.Tables)
                {
                    try
                    {
                        int startPos = table.Range.Start;
                        int endPos = table.Range.End;

                        // 获取表格周围的上下文信息
                        string contextBefore = McpToolsHelpers.GetContextBefore(doc, startPos, 30);
                        string tableTitle = McpToolsHelpers.ExtractTableTitle(doc, startPos, 50);

                        // 创建占位符
                        string placeholder;
                        if (format == "html")
                        {
                            placeholder = $"\n[表格{tableIndex}: {table.Rows.Count}行×{table.Columns.Count}列]";
                            if (!string.IsNullOrEmpty(tableTitle)) placeholder += $"标题:{tableTitle.Trim()}|";
                            if (!string.IsNullOrEmpty(contextBefore)) placeholder += $"前文:{contextBefore.Trim()}";
                            placeholder += $"\n";
                        }
                        else
                        {
                            placeholder = $"\n[表格{tableIndex}: {table.Rows.Count}行×{table.Columns.Count}列]";
                            if (!string.IsNullOrEmpty(tableTitle)) placeholder += $"标题:{tableTitle.Trim()}|";
                            if (!string.IsNullOrEmpty(contextBefore)) placeholder += $"前文:{contextBefore.Trim()}";
                            placeholder += $"\n";
                        }

                        // 记录占位符信息：位置、占位符文本、原始元素长度
                        placeholders.Add((startPos, placeholder, endPos - startPos));

                        tableIndex++;
                    }
                    catch (Exception)
                    {
                        tableIndex++;
                    }
                }

                // 收集图片占位符信息
                int imageIndex = 1;
                foreach (Word.InlineShape shape in doc.InlineShapes)
                {
                    if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                    {
                        try
                        {
                            int startPos = shape.Range.Start;
                            int endPos = shape.Range.End;

                            // 获取图片周围的上下文信息
                            string contextBefore = McpToolsHelpers.GetContextBefore(doc, startPos, 30);
                            string imageCaption = McpToolsHelpers.ExtractImageCaption(doc, startPos, 50);

                            // 创建占位符
                            string placeholder;
                            if (format == "html")
                            {
                                placeholder = $"[图片{imageIndex}: {shape.Width}×{shape.Height}px]";
                                if (!string.IsNullOrEmpty(imageCaption)) placeholder += $"说明:{imageCaption.Trim()}|";
                                if (!string.IsNullOrEmpty(contextBefore)) placeholder += $"前文:{contextBefore.Trim()}";
                            }
                            else
                            {
                                placeholder = $"[图片{imageIndex}: {shape.Width}×{shape.Height}px]";
                                if (!string.IsNullOrEmpty(imageCaption)) placeholder += $"说明:{imageCaption.Trim()}|";
                                if (!string.IsNullOrEmpty(contextBefore)) placeholder += $"前文:{contextBefore.Trim()}";
                            }

                            placeholders.Add((startPos, placeholder, endPos - startPos));

                            imageIndex++;
                        }
                        catch (Exception)
                        {
                            imageIndex++;
                        }
                    }
                }

                // 如果没有占位符需要插入，直接返回原始内容
                if (placeholders.Count == 0)
                {
                    return originalContent;
                }

                // 按位置排序占位符（从后往前处理，避免位置偏移）
                placeholders.Sort((a, b) => b.position.CompareTo(a.position));

                var resultBuilder = new StringBuilder(originalContent);

                foreach (var (position, placeholder, originalLength) in placeholders)
                {
                    try
                    {
                        // 将Word位置转换为字符串位置
                        // 这里做一个简化：假设前面的占位符插入不会影响后续位置
                        // 在更精确的实现中，需要维护位置映射表

                        // 查找原始元素在字符串中的大致位置
                        // 这里使用简单的启发式方法
                        int stringPos = ApproximateStringPosition(originalContent, position);

                        if (stringPos >= 0 && stringPos < resultBuilder.Length)
                        {
                                // 检查这个位置是否有表格/图片的特征
                            string context = GetStringContext(resultBuilder.ToString(), stringPos, 20);
                            if (context.Contains("表格") || context.Contains("图片") ||
                                context.Contains("Table") || context.Contains("Figure"))
                            {
                                // 在元素位置插入占位符，替换原有的内容
                                int replaceLength = Math.Min(originalLength, 50); // 最多替换50个字符
                                resultBuilder.Remove(stringPos, replaceLength);
                                resultBuilder.Insert(stringPos, placeholder);
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                string result = resultBuilder.ToString();

                return result;
            }
            catch (Exception)
            {
                return originalContent;
            }
        }

        /// <summary>
        /// 将Word文档位置近似转换为字符串位置
        /// </summary>
        private static int ApproximateStringPosition(string content, int wordPosition)
        {
            // 这是一个简化的映射方法
            // 实际的Word位置到字符串位置映射非常复杂
            // 这里使用比例估算
            if (string.IsNullOrEmpty(content)) return -1;

            double ratio = (double)wordPosition / 1000; // 假设文档长度约为1000字符
            int estimatedPos = (int)(content.Length * ratio);

            // 确保位置在有效范围内
            return Math.Max(0, Math.Min(estimatedPos, content.Length - 1));
        }

        /// <summary>
        /// 获取字符串中指定位置周围的上下文
        /// </summary>
        private static string GetStringContext(string content, int position, int contextLength)
        {
            try
            {
                int start = Math.Max(0, position - contextLength);
                int end = Math.Min(content.Length, position + contextLength);
                return content.Substring(start, end - start);
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 建立Word位置到字符串位置的映射
        /// </summary>
        private static List<(int WordPos, int StringPos)> BuildPositionMapping(Word.Document doc, string content)
        {
            var mapping = new List<(int WordPos, int StringPos)>();
            try
            {
                int stringPos = 0;
                Word.Range fullRange = doc.Content;

                for (int i = 1; i <= fullRange.Characters.Count && stringPos < content.Length; i++)
                {
                    try
                    {
                        Word.Range charRange = fullRange.Characters[i];
                        string charText = charRange.Text;

                        if (!string.IsNullOrEmpty(charText))
                        {
                            // 记录Word位置到字符串位置的映射
                            mapping.Add((charRange.Start, stringPos));

                            // 移动字符串位置
                            int charLength = charText.Length;
                            if (stringPos + charLength <= content.Length)
                            {
                                // 检查字符是否匹配（处理可能的编码差异）
                                string contentChar = content.Substring(stringPos, charLength);
                                if (contentChar == charText)
                                {
                                    stringPos += charLength;
                                }
                                else
                                {
                                    // 如果不匹配，尝试找到下一个匹配位置
                                    int nextMatch = content.IndexOf(charText, stringPos);
                                    if (nextMatch >= 0)
                                    {
                                        stringPos = nextMatch + charLength;
                                    }
                                    else
                                    {
                                        stringPos += charLength; // 继续前进
                                    }
                                }
                            }
                            else
                            {
                                break; // 超出内容长度
                            }
                        }
                    }
                    catch (Exception)
                    {
                        break;
                    }
                }

                return mapping;
            }
            catch (Exception)
            {
                return mapping;
            }
        }

        /// <summary>
        /// 过滤临时内容（使用历史删除记录进行过滤）
        /// </summary>
        private static string FilterTempContent(Word.Document doc, string originalContent, string format)
        {
            try
            {
                string filteredContent = originalContent;

                // 使用最终删除记录进行过滤（考虑按钮偏移后的实际范围）
                var finalDeleteRanges = DocumentState.FinalDeleteRanges;
                var finalInsertRanges = DocumentState.FinalInsertRanges;

                // 如果没有最终范围，则使用历史范围作为后备
                if (finalDeleteRanges.Count == 0 && finalInsertRanges.Count == 0)
                {
                    finalDeleteRanges = DocumentState.HistoricalDeleteRanges;
                    finalInsertRanges = DocumentState.HistoricalInsertRanges;
                }

                if (finalDeleteRanges.Count > 0)
                {
                    // 创建位置映射：Word位置 -> 字符串位置
                    var positionMap = BuildPositionMapping(doc, originalContent);

                    // 收集需要移除的文本片段（按位置从后往前处理，避免位置偏移）
                    var textFragmentsToRemove = new List<(int stringStart, int length, string text)>();

                    foreach (var range in finalDeleteRanges)
                    {
                        int wordStart = range.Start;
                        int wordEnd = range.End;

                        try
                        {
                            // 将Word位置转换为字符串位置（新规则：[start,end]包含结束位置）
                            int stringStart = -1;
                            int stringEnd = -1;

                            // 查找精确的字符串位置
                            foreach (var mapping in positionMap)
                            {
                                if (mapping.WordPos == wordStart)
                                {
                                    stringStart = mapping.StringPos;
                                }
                                if (mapping.WordPos == wordEnd)
                                {
                                    stringEnd = mapping.StringPos;
                                }
                            }

                            // 如果没找到精确匹配，使用最接近的匹配
                            if (stringStart == -1)
                            {
                                foreach (var mapping in positionMap)
                                {
                                    if (mapping.WordPos <= wordStart)
                                    {
                                        stringStart = mapping.StringPos;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }

                            if (stringEnd == -1)
                            {
                                foreach (var mapping in positionMap)
                                {
                                    if (mapping.WordPos <= wordEnd)
                                    {
                                        stringEnd = mapping.StringPos;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }

                            // 新规则：范围[start,end]包含结束位置，所以需要包含wordEnd对应的字符
                            // 如果存在wordEnd的映射，结束位置应该是该字符后的位置
                            if (stringEnd >= 0 && stringEnd < originalContent.Length)
                            {
                                var nextMappings = positionMap.Where(m => m.WordPos > wordEnd).ToList();
                                if (nextMappings.Any())
                                {
                                    stringEnd = nextMappings.First().StringPos;
                                }
                                else
                                {
                                    // 如果没有更多的映射，意味着wordEnd是最后一个字符，扩展到内容末尾
                                    stringEnd = originalContent.Length;
                                }
                            }

                            if (stringStart >= 0 && stringEnd > stringStart && stringEnd <= originalContent.Length)
                            {
                                int length = stringEnd - stringStart;
                                string textToRemove = originalContent.Substring(stringStart, length);
                                textFragmentsToRemove.Add((stringStart, length, textToRemove));
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }

                    // 按位置从后往前排序，确保移除时位置正确
                    textFragmentsToRemove.Sort((a, b) => b.stringStart.CompareTo(a.stringStart));

                    // 从后往前移除文本片段
                    foreach (var (stringStart, length, text) in textFragmentsToRemove)
                    {
                        try
                        {
                            if (stringStart + length <= filteredContent.Length)
                            {
                                filteredContent = filteredContent.Remove(stringStart, length);
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }

                // 使用格式检查过滤接受和丢弃按钮文本
                string[] buttonPatterns = new[] { " 接受 ", " 丢弃 " };
                foreach (string pattern in buttonPatterns)
                {
                    if (filteredContent.Contains(pattern))
                    {
                        filteredContent = filteredContent.Replace(pattern, "");
                    }
                }

                // 额外清理：移除所有带有接受/丢弃按钮背景色的文本
                try
                {
                    Word.WdColor acceptBtnBg = ConfigManager.DocumentColorHelper.AcceptButtonBackground;
                    Word.WdColor rejectBtnBg = ConfigManager.DocumentColorHelper.RejectButtonBackground;

                    Word.Range fullRange = doc.Content;
                    var buttonTextsToRemove = new List<string>();

                    for (int i = 1; i <= fullRange.Characters.Count; i++)
                    {
                        Word.Range charRange = fullRange.Characters[i];
                        Word.WdColor bgColor = charRange.Shading.BackgroundPatternColor;
                        string charText = charRange.Text;

                        // 检测接受或丢弃按钮背景色
                        if ((bgColor == acceptBtnBg || bgColor == rejectBtnBg) && charText.Trim().Length > 0)
                        {
                            if (!buttonTextsToRemove.Contains(charText))
                            {
                                buttonTextsToRemove.Add(charText);
                            }
                        }
                    }

                    // 移除找到的按钮文本
                    foreach (string buttonText in buttonTextsToRemove)
                    {
                        if (filteredContent.Contains(buttonText))
                        {
                            filteredContent = filteredContent.Replace(buttonText, "");
                        }
                    }
                }
                catch (Exception)
                {
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 过滤后内容前200字符: {filteredContent.Substring(0, Math.Min(200, filteredContent.Length))}");

                return filteredContent;
            }
            catch (Exception)
            {
                return originalContent; // 过滤失败时返回原内容
            }
        }
    }
}
