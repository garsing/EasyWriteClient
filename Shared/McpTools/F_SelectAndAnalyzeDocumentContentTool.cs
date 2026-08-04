using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 文档内容选择和分析工具
    /// 在Word文档中选中指定的内容范围，并返回选中内容的详细结构信息
    /// </summary>
    public static class F_SelectAndAnalyzeDocumentContentTool
    {
        /// <summary>
        /// 注册文档内容选择工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_select_and_analyze_document_content"] = async (args) =>
            {
                await Task.CompletedTask; // 确保异步执行

                try
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] select_and_analyze_document_content工具开始执行");

                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序实例不可用" };
                    }

                    dynamic wordApp = wordApplication;

                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    string selectionType = args.ContainsKey("selection_type") ? args["selection_type"]?.ToString() : "";
                    bool includeTableStructure = !args.ContainsKey("include_table_structure") || Convert.ToBoolean(args["include_table_structure"]);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 选择类型: {selectionType}, 包含表格结构: {includeTableStructure}");

                    Word.Range selectedRange = null;
                    string selectionDescription = "";

                    switch (selectionType.ToLower())
                    {
                        case "paragraph_by_content":
                            // 按段落内容特征选择
                            string paraStart = args.ContainsKey("paragraph_start") ? args["paragraph_start"]?.ToString() : "";
                            string paraEnd = args.ContainsKey("paragraph_end") ? args["paragraph_end"]?.ToString() : "";

                            if (string.IsNullOrEmpty(paraStart))
                            {
                                return new ToolResult { Success = false, Error = "选择类型为'paragraph_by_content'时，必须提供paragraph_start参数" };
                            }

                            selectedRange = SelectParagraphByContent(wordApp.ActiveDocument, paraStart, paraEnd);
                            if (selectedRange == null)
                            {
                                return new ToolResult { Success = false, Error = $"未能找到包含开头内容'{paraStart}'的段落" };
                            }

                            selectionDescription = $"段落：开头'{paraStart}'";
                            if (!string.IsNullOrEmpty(paraEnd))
                            {
                                // 检查结尾是否找到并使用了
                                Word.Range endRange = McpToolsHelpers.FindText(wordApp.ActiveDocument.Content, paraEnd);
                                if (endRange != null && endRange.Start > McpToolsHelpers.FindText(wordApp.ActiveDocument.Content, paraStart).Start)
                                {
                                    selectionDescription += $"结尾'{paraEnd}'";
                                }
                                else
                                {
                                    selectionDescription += $"（结尾'{paraEnd}'未找到，仅选择段落开头部分）";
                                }
                            }
                            break;

                        case "table_by_index":
                            // 按表格序号选择
                            if (!args.ContainsKey("table_index"))
                            {
                                return new ToolResult { Success = false, Error = "选择类型为'table_by_index'时，必须提供table_index参数" };
                            }

                            int tableIdx = Convert.ToInt32(args["table_index"]);
                            selectedRange = SelectTableByIndex(wordApp.ActiveDocument, tableIdx);
                            selectionDescription = $"表格 {tableIdx}";
                            break;

                        case "table_by_context":
                            // 按表格上下文选择
                            string tableBefore = args.ContainsKey("table_context_before") ? args["table_context_before"]?.ToString() : "";
                            string tableAfter = args.ContainsKey("table_context_after") ? args["table_context_after"]?.ToString() : "";

                            if (string.IsNullOrEmpty(tableBefore))
                            {
                                return new ToolResult { Success = false, Error = "选择类型为'table_by_context'时，必须提供table_context_before参数" };
                            }

                            selectedRange = SelectTableByContext(wordApp.ActiveDocument, tableBefore, tableAfter);
                            selectionDescription = $"表格：上下文'{tableBefore}'" + (string.IsNullOrEmpty(tableAfter) ? "" : $"到'{tableAfter}'");
                            break;

                        case "image_by_context":
                            // 按图片上下文选择
                            string imageBefore = args.ContainsKey("image_context_before") ? args["image_context_before"]?.ToString() : "";
                            string imageAfter = args.ContainsKey("image_context_after") ? args["image_context_after"]?.ToString() : "";

                            if (string.IsNullOrEmpty(imageBefore))
                            {
                                return new ToolResult { Success = false, Error = "选择类型为'image_by_context'时，必须提供image_context_before参数" };
                            }

                            selectedRange = SelectImageByContext(wordApp.ActiveDocument, imageBefore, imageAfter);
                            selectionDescription = $"图片：上下文'{imageBefore}'" + (string.IsNullOrEmpty(imageAfter) ? "" : $"到'{imageAfter}'");
                            break;

                        default:
                            return new ToolResult { Success = false, Error = $"不支持的选择类型：{selectionType}" };
                    }

                    if (selectedRange == null)
                    {
                        return new ToolResult { Success = false, Error = $"未能找到指定的内容：{selectionDescription}" };
                    }

                    // 激活Word窗口并执行选中
                    bool windowActivated = false;
                    try
                    {
                        wordApp.Activate();
                        if (wordApp.ActiveWindow != null)
                        {
                            wordApp.ActiveWindow.Activate();
                            windowActivated = true;
                        }
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Word窗口激活成功");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 激活Word窗口失败: {ex.Message}");
                    }

                    // 执行选中操作
                    try
                    {
                        selectedRange.Select();
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 选中操作执行成功");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 选中操作失败: {ex.Message}");
                        return new ToolResult { Success = false, Error = $"选中操作失败: {ex.Message}" };
                    }

                    // 获取选择结果信息
                    string previewText = selectedRange.Text?.Length > 100
                        ? selectedRange.Text.Substring(0, 100) + "..."
                        : selectedRange.Text ?? "";

                    // 准备返回数据（使用字典以便动态添加字段）
                    var resultData = new Dictionary<string, object>
                    {
                        ["selection_type"] = selectionType,
                        ["selection_description"] = selectionDescription,
                        ["start_position"] = selectedRange.Start,
                        ["end_position"] = selectedRange.End,
                        ["selected_length"] = selectedRange.End - selectedRange.Start,
                        ["preview_text"] = previewText,
                        ["document_name"] = wordApp.ActiveDocument.Name ?? "未命名文档",
                        ["window_activated"] = windowActivated,
                        ["selection_completed"] = true,
                        ["message"] = $"已成功选中{selectionDescription}，Word窗口已激活并显示选中状态。您现在可以进行编辑、复制或其他操作。"
                    };

                    // 如果选中了表格且需要包含结构信息，则添加表格结构
                    if ((selectionType == "table_by_index" || selectionType == "table_by_context") && includeTableStructure)
                    {
                        try
                        {
                            Word.Table selectedTable = null;

                            if (selectionType == "table_by_index")
                            {
                                int tableIdx = Convert.ToInt32(args["table_index"]);
                                if (tableIdx > 0 && tableIdx <= wordApp.ActiveDocument.Tables.Count)
                                {
                                    selectedTable = wordApp.ActiveDocument.Tables[tableIdx];
                                }
                            }
                            else if (selectionType == "table_by_context")
                            {
                                // 通过范围找到对应的表格
                                foreach (Word.Table table in wordApp.ActiveDocument.Tables)
                                {
                                    if (table.Range.Start >= selectedRange.Start && table.Range.End <= selectedRange.End)
                                    {
                                        selectedTable = table;
                                        break;
                                    }
                                }
                            }

                            if (selectedTable != null)
                            {
                                var tableStructure = GetTableDetailedStructure(selectedTable);
                                resultData["table_structure"] = tableStructure;
                                resultData["message"] = $"已成功选中{selectionDescription}，Word窗口已激活并显示选中状态。已返回表格的详细结构信息。";
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取表格结构失败: {ex.Message}");
                        }
                    }

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = resultData
                    };

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] select_and_analyze_document_content工具执行成功，已选中：{selectionDescription}");
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] select_and_analyze_document_content工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"文档内容选择和分析失败: {ex.Message}" };
                }
            };
        }

        /// <summary>
        /// 按段落内容特征选择段落
        /// </summary>
        private static Word.Range SelectParagraphByContent(Word.Document doc, string startContent, string endContent = null)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 查找开头'{startContent}'");

                // 查找开头内容
                Word.Range startRange = McpToolsHelpers.FindText(doc.Content, startContent);
                if (startRange == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 开头内容'{startContent}'未找到");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 找到开头内容，位置: {startRange.Start}");

                // 获取开头内容所在的段落
                Word.Paragraph startParagraph = startRange.Paragraphs[1];
                Word.Range paragraphRange = startParagraph.Range;

                System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 段落范围: {paragraphRange.Start}-{paragraphRange.End}");

                // 如果没有提供结尾内容，直接返回整个段落
                if (string.IsNullOrEmpty(endContent))
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 未提供结尾内容，返回整个段落");
                    return paragraphRange;
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 查找结尾'{endContent}'");

                // 如果提供了结尾内容，尝试找到结尾并扩展选择范围
                Word.Range endRange = McpToolsHelpers.FindText(doc.Content, endContent);
                if (endRange != null && endRange.Start > startRange.Start)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 找到结尾内容，位置: {endRange.Start}");

                    // 创建从开头到结尾的范围
                    Word.Range extendedRange = doc.Range(startRange.Start, endRange.End);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 返回扩展范围: {extendedRange.Start}-{extendedRange.End}");
                    return extendedRange;
                }

                // 如果找不到结尾内容，返回整个段落
                System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 结尾内容'{endContent}'未找到，返回整个段落");
                return paragraphRange;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] SelectParagraphByContent: 发生异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 按表格序号选择表格
        /// </summary>
        private static Word.Range SelectTableByIndex(Word.Document doc, int tableIndex)
        {
            try
            {
                if (tableIndex <= 0 || tableIndex > doc.Tables.Count)
                {
                    return null;
                }

                return doc.Tables[tableIndex].Range;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 按表格上下文选择表格
        /// </summary>
        private static Word.Range SelectTableByContext(Word.Document doc, string contextBefore, string contextAfter = null)
        {
            try
            {
                // 查找前上下文
                Word.Range beforeRange = McpToolsHelpers.FindText(doc.Content, contextBefore);
                if (beforeRange == null) return null;

                // 从前上下文位置开始，向后查找第一个表格
                foreach (Word.Table table in doc.Tables)
                {
                    if (table.Range.Start >= beforeRange.End)
                    {
                        // 如果提供了后上下文，验证表格是否在正确位置
                        if (!string.IsNullOrEmpty(contextAfter))
                        {
                            Word.Range afterRange = McpToolsHelpers.FindText(doc.Content, contextAfter);
                            if (afterRange != null && table.Range.End <= afterRange.Start)
                            {
                                return table.Range;
                            }
                        }
                        else
                        {
                            // 没有后上下文，直接返回找到的表格
                            return table.Range;
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 按图片上下文选择图片
        /// </summary>
        private static Word.Range SelectImageByContext(Word.Document doc, string contextBefore, string contextAfter = null)
        {
            try
            {
                // 查找前上下文
                Word.Range beforeRange = McpToolsHelpers.FindText(doc.Content, contextBefore);
                if (beforeRange == null) return null;

                // 从前上下文位置开始，向后查找第一个图片
                foreach (Word.InlineShape shape in doc.InlineShapes)
                {
                    if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture &&
                        shape.Range.Start >= beforeRange.End)
                    {
                        // 如果提供了后上下文，验证图片是否在正确位置
                        if (!string.IsNullOrEmpty(contextAfter))
                        {
                            Word.Range afterRange = McpToolsHelpers.FindText(doc.Content, contextAfter);
                            if (afterRange != null && shape.Range.End <= afterRange.Start)
                            {
                                return shape.Range;
                            }
                        }
                        else
                        {
                            // 没有后上下文，直接返回找到的图片
                            return shape.Range;
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取表格的详细内部结构信息
        /// </summary>
        private static object GetTableDetailedStructure(Word.Table table)
        {
            try
            {
                var rows = new List<object>();

                for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
                {
                    var cells = new List<object>();
                    Word.Row row = table.Rows[rowIndex];

                    for (int colIndex = 1; colIndex <= table.Columns.Count; colIndex++)
                    {
                        try
                        {
                            Word.Cell cell = table.Cell(rowIndex, colIndex);
                            string cellText = cell.Range.Text?.Trim();

                            // 移除Word表格单元格中的特殊字符
                            if (!string.IsNullOrEmpty(cellText))
                            {
                                // 移除末尾的段落标记和制表符
                                cellText = cellText.Replace("\r", "").Replace("\a", "").Replace("\t", "").Trim();
                            }

                            cells.Add(new
                            {
                                row_index = rowIndex,
                                column_index = colIndex,
                                content = cellText ?? "",
                                is_merged = false // 暂时设为false，后续可以通过其他方式检测合并单元格
                            });
                        }
                        catch (Exception cellEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取单元格({rowIndex},{colIndex})失败: {cellEx.Message}");
                            cells.Add(new
                            {
                                row_index = rowIndex,
                                column_index = colIndex,
                                content = "",
                                is_merged = false
                            });
                        }
                    }

                    rows.Add(new
                    {
                        row_index = rowIndex,
                        cells = cells
                    });
                }

                return new
                {
                    rows_count = table.Rows.Count,
                    columns_count = table.Columns.Count,
                    rows = rows
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取表格详细结构失败: {ex.Message}");
                return new
                {
                    rows_count = table.Rows.Count,
                    columns_count = table.Columns.Count,
                    rows = new List<object>(),
                    error = $"获取表格结构失败: {ex.Message}"
                };
            }
        }
    }
}
