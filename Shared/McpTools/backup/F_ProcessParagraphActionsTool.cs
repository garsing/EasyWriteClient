using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 处理段落操作工具
    /// 支持通过段落号对Word文档进行替换、删除、插入等操作
    /// </summary>
    public static class F_ProcessParagraphActionsTool
    {
        /// <summary>
        /// 注册处理段落操作工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            List<McpTool> availableTools,
            object wordApplication)
        {
            var tool = new McpTool
            {
                name = "F_process_paragraph_actions",
                alias = "处理段落操作",
                description = "通过文本定位对Word文档进行批量段落操作，支持替换、删除、插入段落内容。⚠️重要：本工具是以段落为级别的修改工具，所有操作都是针对整个段落进行的。必须确认源文档选中的一定是一个段落，修改的新内容也必须是一整个段落。本工具无法只修改一句话，若想修改段落中的一句话，必须整个段落进行修改。调用格式：{\"action\": [{\"type\": \"replace\", \"detail\": {\"original\": {\"start\": \"开头文本\"(必需), \"end\": \"结尾文本\"(可选)}, \"new\": \"新内容\"(必需，必须是完整段落)}}, {\"type\": \"delete\", \"detail\": {\"deletion\": {\"start\": \"开头文本\"(必需), \"end\": \"结尾文本\"(可选)}}}, {\"type\": \"insert\", \"detail\": {\"pre\": \"上文内容\"(必需), \"insert\": \"插入内容\"(必需，必须是完整段落), \"post\": \"下文内容\"(必需)}}]}。替换显示原文本(删除线)+新文本(黄色背景)，删除保留原文本(删除线)，插入显示新文本(黄色背景)。注意：replace和delete操作至少需要提供start文本来定位段落，end文本可选；如果同时提供start和end，会精确定位从start到end的文本范围所在的段落；如果只提供start，会定位包含start文本的段落。如果操作失败请先用F_get_document_structure检查文档状态",
                usage = @"批量修改文档段落内容。⚠️重要：本工具是以段落为级别的修改工具，所有操作都是针对整个段落进行的。

⚠️关键限制：
- 必须确认源文档选中的一定是一个段落（通过start/end文本定位到的段落）
- 修改的新内容也必须是一整个段落
- 本工具无法只修改一句话
- 若想修改段落中的一句话，必须整个段落进行修改

支持三种操作：

1) replace-替换段落：
   - 必需字段：original.start（原文开头文本，用于定位段落）、new（新内容，必须是完整段落）
   - 可选字段：original.end（原文结尾文本，如果提供会精确定位从start到end的文本范围所在的段落）
   - 如果同时提供start和end：精确定位从start到end的文本范围所在的段落，然后替换整个段落
   - 如果只提供start：定位包含start文本的段落，然后替换整个段落
   - ⚠️注意：new字段必须是完整的段落内容，不能只是一句话
   - 显示效果：原文本(删除线)+新文本(黄色背景)

2) delete-删除段落：
   - 必需字段：deletion.start（要删除文本的开头，用于定位段落）
   - 可选字段：deletion.end（要删除文本的结尾，如果提供会精确定位从start到end的文本范围所在的段落）
   - 如果同时提供start和end：精确定位从start到end的文本范围所在的段落，然后删除整个段落
   - 如果只提供start：定位包含start文本的段落，然后删除整个段落
   - ⚠️注意：删除操作会删除整个段落，不能只删除段落中的一句话
   - 显示效果：保留原文本并添加删除线标记

3) insert-插入段落：
   - 必需字段：pre（上文内容，用于定位插入位置）、insert（要插入的内容，必须是完整段落）、post（下文内容，用于定位插入位置）
   - 通过pre和post定位插入位置，在定位位置之后插入新文本
   - ⚠️注意：insert字段必须是完整的段落内容，不能只是一句话
   - 显示效果：新文本(黄色背景)

系统自动添加换行符，确保段落不会被合并",
                input_schema = new McpTool.ToolInputSchema
                {
                    Type = "object",
                    Properties = new Dictionary<string, object>
                    {
                        ["action"] = new
                        {
                            type = "array",
                            description = "要执行的操作列表",
                            items = new
                            {
                                type = "object",
                                properties = new Dictionary<string, object>
                                {
                                    ["type"] = new
                                    {
                                        type = "string",
                                        @enum = new[] { "replace", "delete", "insert" },
                                        description = "操作类型：replace替换、delete删除、insert插入"
                                    },
                                    ["detail"] = new
                                    {
                                        type = "object",
                                        description = "操作详情，根据操作类型不同而不同"
                                    }
                                },
                                required = new[] { "type", "detail" }
                            }
                        }
                    },
                    required = new List<string> { "action" }
                }
            };

            toolRegistry["F_process_paragraph_actions"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序实例不可用" };
                    }

                    dynamic wordApp = wordApplication;

                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作前文档状态: 段落总数={wordApp.ActiveDocument.Paragraphs.Count}");

                    // 第一步：清除原有的操作块（格式和按钮）
                    System.Diagnostics.Debug.WriteLine("[第一步] 清除原有的操作块（格式和按钮）");
                    ClearOperationBlocks(wordApp.ActiveDocument);
                    System.Diagnostics.Debug.WriteLine("[第一步完成] 清除操作块完成");

                    // 打印操作前所有段落内容用于诊断
                    var paragraphsBefore = new List<string>();
                    for (int i = 1; i <= wordApp.ActiveDocument.Paragraphs.Count; i++)
                    {
                        var para = wordApp.ActiveDocument.Paragraphs[i];
                        string content = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作前段落{i}: 范围({para.Range.Start}-{para.Range.End}), 内容='{content.Substring(0, Math.Min(30, content.Length))}'");

                        // 收集所有段落用于跟踪（包括空段落）
                        paragraphsBefore.Add(content);
                    }

                    // 第二步：初始化段落跟踪（如果还没有快照的话）
                    System.Diagnostics.Debug.WriteLine("[第二步] 检查段落跟踪状态");
                    var currentSnapshot = DocumentState.CurrentSnapshot;
                    if (currentSnapshot == null || currentSnapshot.Count == 0)
                    {
                        // 没有快照，需要初始化
                        System.Diagnostics.Debug.WriteLine("[第二步] 初始化段落跟踪（首次调用）");
                        DocumentState.InitializeParagraphTracking(paragraphsBefore);
                    }
                    else
                    {
                        // 已有快照，基于最新快照继续（不重新初始化）
                        System.Diagnostics.Debug.WriteLine($"[第二步] 已有快照，基于最新快照继续（快照数量: {DocumentState.Snapshots.Count}）");
                    }
                    System.Diagnostics.Debug.WriteLine("[第二步完成]");

                    if (!args.ContainsKey("action"))
                    {
                        return new ToolResult { Success = false, Error = "缺少必需参数：action" };
                    }

                    List<ParagraphAction> actions = ParseActions(args["action"], wordApp.ActiveDocument);
                    
                    // 第三步：生成新的快照（不操作文档）
                    System.Diagnostics.Debug.WriteLine("[第三步] 生成新的快照（不操作文档）");
                    string errorMessage = await ProcessActionsForSnapshot(actions, wordApp.ActiveDocument);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        // 出现异常，恢复按钮和格式：运行第六步根据display_hunk修改格式
                        System.Diagnostics.Debug.WriteLine("[异常恢复] 出现异常，恢复按钮和格式");
                        var restoreDisplayHunk = DocumentState.DisplayHunk;
                        if (restoreDisplayHunk != null && restoreDisplayHunk.Count > 0)
                        {
                            // 将display_hunk转换为hunk格式（List<(string header, List<string> lines)>）
                            var restoreHunk = new List<(string header, List<string> lines)>();
                            foreach (var (header, lines) in restoreDisplayHunk)
                            {
                                restoreHunk.Add((header, lines.ToList()));
                            }
                            
                            // 获取targetFullDisplayContent
                            var restoreTargetFullDisplayContent = DocumentState.GetTargetFullDisplayContent(restoreHunk);
                            
                            // 运行第六步：根据全量hunk修改格式
                            System.Diagnostics.Debug.WriteLine("[异常恢复] 执行第六步：根据display_hunk修改格式");
                            foreach (var (hunkHeader, hunkLines) in restoreHunk)
                            {
                                ApplyHunkFormatting(wordApp.ActiveDocument, hunkLines, restoreTargetFullDisplayContent);
                            }
                            System.Diagnostics.Debug.WriteLine("[异常恢复] 格式和按钮已恢复");
                        }
                        
                        return new ToolResult { Success = false, Error = errorMessage };
                    }
                    System.Diagnostics.Debug.WriteLine("[第三步完成]");

                    // 第四步：生成全量hunk
                    System.Diagnostics.Debug.WriteLine("[第四步] 生成全量hunk差异");
                    var hunk = DocumentState.CalculateAndPrintHunk();
                    System.Diagnostics.Debug.WriteLine($"[第四步完成] 共生成 {hunk.Count} 个全量hunk（包含所有段落）");

                    // 第五步：根据全量hunk执行文档插入操作（只处理+号的段落）
                    System.Diagnostics.Debug.WriteLine("[第五步] 根据全量hunk执行文档插入操作");
                    // 获取起点（originalFullDisplayContent）：从当前的display_hunk直接获取（全量形式）
                    var originalFullDisplayContent = DocumentState.GetOriginalFullDisplayContent();
                    System.Diagnostics.Debug.WriteLine($"[第五步] originalFullDisplayContent（起点）: {string.Join(",", originalFullDisplayContent)}");
                    
                    // 获取终点（targetFullDisplayContent）：从新的全量hunk直接获取
                    var targetFullDisplayContent = DocumentState.GetTargetFullDisplayContent(hunk);
                    System.Diagnostics.Debug.WriteLine($"[第五步] targetFullDisplayContent（终点）: {string.Join(",", targetFullDisplayContent)}");
                    
                    // 对比起点和终点，找出需要插入的段落
                    await ApplyHunkInsertions(wordApp.ActiveDocument, hunk, originalFullDisplayContent, targetFullDisplayContent);
                    System.Diagnostics.Debug.WriteLine("[第五步完成]");

                    // 第五步-2：更新display_hunk
                    System.Diagnostics.Debug.WriteLine("[第五步-2] 更新display_hunk");
                    // 直接存储全量hunk列表（包含所有段落）
                    DocumentState.UpdateDisplayHunk(hunk);

                    // 打印更新后的display_hunk用于调试
                    var displayHunk = DocumentState.DisplayHunk;
                    System.Diagnostics.Debug.WriteLine($"[第五步-2] display_hunk更新完成，当前包含 {displayHunk.Count} 个全量hunk:");
                    foreach (var (header, lines) in displayHunk)
                    {
                        System.Diagnostics.Debug.WriteLine($"  全量Hunk (header为空): 包含 {lines.Count} 行段落");
                        foreach (var line in lines)
                        {
                            // 输出所有行，包括空行
                            System.Diagnostics.Debug.WriteLine($"    {line}");
                        }
                    }
                    System.Diagnostics.Debug.WriteLine("[第五步-2完成]");

                    // 第六步：根据全量hunk修改格式
                    System.Diagnostics.Debug.WriteLine("[第六步] 根据全量hunk修改格式");
                    foreach (var (hunkHeader, hunkLines) in hunk)
                    {
                        // 全量hunk：直接使用hunkLines（包含所有段落）
                        ApplyHunkFormatting(wordApp.ActiveDocument, hunkLines, targetFullDisplayContent);
                    }
                    System.Diagnostics.Debug.WriteLine("[第六步完成]");

                    // 为了兼容性，仍然返回操作结果
                    var results = actions.Select((ParagraphAction action, int index) => new ActionResult
                    {
                        action_index = index,
                        type = action.Type,
                        target_index = action.TargetIndex,
                        success = true,
                        message = "操作成功"
                    }).ToList();

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            document_name = wordApp.ActiveDocument.Name ?? "未命名文档",
                            total_actions = actions.Count,
                            executed_actions = results.Count,
                            results = results
                        }
                    };

                    return result;
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"处理段落操作失败: {ex.Message}" };
                }
            };

            availableTools.Add(tool);
        }

        /// <summary>
        /// 解析操作列表（新格式：存储start/end文本，稍后在ProcessActionsForSnapshot中查找段落序号）
        /// </summary>
        private static List<ParagraphAction> ParseActions(object actionsObj, Word.Document doc)
        {
            var actions = new List<ParagraphAction>();

            try
            {
                if (actionsObj is System.Collections.IEnumerable actionsList)
                {
                    foreach (var actionObj in actionsList)
                    {
                        var actionDict = actionObj as Dictionary<string, object>;
                        if (actionDict == null || !actionDict.ContainsKey("type"))
                            continue;

                        var action = new ParagraphAction();
                        action.Type = actionDict["type"]?.ToString();
                        
                        var detail = actionDict.ContainsKey("detail") ? actionDict["detail"] as Dictionary<string, object> : null;
                        if (detail == null)
                            continue;

                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                                // 解析replace操作：original.start, original.end, new
                                object original = detail.ContainsKey("original") ? detail["original"] : null;
                                if (original is Dictionary<string, object> originalDict)
                                {
                                    action.StartText = originalDict.ContainsKey("start") ? originalDict["start"]?.ToString() : null;
                                    action.EndText = originalDict.ContainsKey("end") ? originalDict["end"]?.ToString() : null;
                                }
                                
                                // 解析新内容
                                if (detail.ContainsKey("new"))
                                {
                                    action.Content = detail["new"]?.ToString();
                                }
                                break;

                            case "delete":
                                // 解析delete操作：deletion.start, deletion.end
                                object deletion = detail.ContainsKey("deletion") ? detail["deletion"] : null;
                                if (deletion is Dictionary<string, object> deletionDict)
                                {
                                    action.StartText = deletionDict.ContainsKey("start") ? deletionDict["start"]?.ToString() : null;
                                    action.EndText = deletionDict.ContainsKey("end") ? deletionDict["end"]?.ToString() : null;
                                }
                                break;

                            case "insert":
                                // 解析insert操作：pre, insert, post
                                action.PreText = detail.ContainsKey("pre") ? detail["pre"]?.ToString() : null;
                                action.PostText = detail.ContainsKey("post") ? detail["post"]?.ToString() : null;
                                action.Position = "after"; // 默认在定位段落之后插入
                                
                                // 解析插入内容
                                if (detail.ContainsKey("insert"))
                                {
                                    action.Content = detail["insert"]?.ToString();
                                }
                                break;
                        }

                        // 验证必要字段
                        bool isValid = false;
                        if (action.Type?.ToLower() == "replace")
                        {
                            // replace操作：至少需要start文本，end文本可选；还需要new内容
                            isValid = !string.IsNullOrEmpty(action.StartText) && !string.IsNullOrEmpty(action.Content);
                        }
                        else if (action.Type?.ToLower() == "delete")
                        {
                            // delete操作：至少需要start文本，end文本可选
                            isValid = !string.IsNullOrEmpty(action.StartText);
                        }
                        else if (action.Type?.ToLower() == "insert")
                        {
                            isValid = !string.IsNullOrEmpty(action.PreText) && !string.IsNullOrEmpty(action.PostText) && !string.IsNullOrEmpty(action.Content);
                        }

                        if (isValid)
                        {
                            actions.Add(action);
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 解析操作失败：缺少必要字段，操作类型={action.Type}, StartText={action.StartText}, EndText={action.EndText}, Content={action.Content}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 解析操作失败: {ex.Message}");
            }

            return actions;
        }

        /// <summary>
        /// 通过文本范围找到段落序号
        /// </summary>
        private static int FindParagraphIndexByRange(Word.Document doc, Word.Range range)
        {
            try
            {
                if (range == null)
                    return 0;

                // 获取范围所在的段落
                Word.Paragraph para = range.Paragraphs[1];
                
                // 遍历文档段落找到对应的索引
                for (int i = 1; i <= doc.Paragraphs.Count; i++)
                {
                    if (doc.Paragraphs[i].Range.Start == para.Range.Start)
                    {
                        return i;
                    }
                }
                
                return 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 查找段落序号失败: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 处理段落操作
        /// </summary>
        private static async Task<List<ActionResult>> ProcessActions(Word.Document doc, List<ParagraphAction> actions)
        {
            var results = new List<ActionResult>();

            try
            {
                // 按段落索引从大到小排序处理，避免索引变化影响
                var sortedActions = actions
                    .Select((action, originalIndex) => new { Action = action, OriginalIndex = originalIndex })
                    .OrderByDescending(x => x.Action.TargetIndex)
                    .ToList();

                foreach (var item in sortedActions)
                {
                    var action = item.Action;
                    var result = new ActionResult
                    {
                        action_index = item.OriginalIndex,
                        type = action.Type,
                        target_index = action.TargetIndex
                    };

                    try
                    {
                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                                System.Diagnostics.Debug.WriteLine($"[操作步骤] 执行替换操作：段落{action.TargetIndex}");
                                result.success = ReplaceParagraph(doc, action.TargetIndex, action.Content);
                                if (result.success)
                                {
                                    DocumentState.ProcessParagraphReplace(action.TargetIndex, action.Content);
                                }
                                result.message = result.success ? "替换成功" : "替换失败：段落不存在";
                                System.Diagnostics.Debug.WriteLine($"[操作结果] {result.message}");
                                break;

                            case "delete":
                                System.Diagnostics.Debug.WriteLine($"[操作步骤] 执行删除操作：段落{action.TargetIndex}");
                                result.success = DeleteParagraph(doc, action.TargetIndex);
                                if (result.success)
                                {
                                    DocumentState.ProcessParagraphDelete(action.TargetIndex);
                                }
                                result.message = result.success ? "删除成功" : "删除失败：段落不存在";
                                System.Diagnostics.Debug.WriteLine($"[操作结果] {result.message}");
                                break;

                            case "insert":
                                System.Diagnostics.Debug.WriteLine($"[操作步骤] 执行插入操作：在段落{action.TargetIndex} {action.Position} 插入内容");
                                result.success = InsertParagraph(doc, action.TargetIndex, action.Content, action.Position);
                                if (result.success)
                                {
                                    DocumentState.ProcessParagraphInsert(action.TargetIndex, action.Position, action.Content);
                                }
                                result.message = result.success ? $"插入成功（{action.Position}）" : "插入失败：目标段落不存在";
                                System.Diagnostics.Debug.WriteLine($"[操作结果] {result.message}");
                                break;

                            default:
                                result.success = false;
                                result.message = $"未知操作类型: {action.Type}";
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.success = false;
                        result.message = $"操作失败: {ex.Message}";
                    }

                    results.Add(result);
                }

                // 按原始顺序返回结果
                results = results.OrderBy(r => r.action_index).ToList();

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 处理操作失败: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// 处理段落操作（只更新快照，不操作文档）
        /// 通过display_hunk和当前快照建立映射关系
        /// 使用两个指针同时遍历display_hunk和当前快照来建立映射
        /// </summary>
        /// <returns>错误信息，如果成功则返回null或空字符串</returns>
        private static async Task<string> ProcessActionsForSnapshot(List<ParagraphAction> actions, Word.Document doc)
        {
            try
            {
                var documentProcessor = new DocumentProcessor(doc.Application);
                
                // 获取当前快照和display_hunk
                var currentSnapshot = DocumentState.CurrentSnapshot.ToList();
                var displayHunk = DocumentState.DisplayHunk;
                
                // 建立display_hunk位置到当前快照位置的映射
                // display_hunk位置（从1开始，对应Word段落位置）-> 当前快照位置（从1开始）
                var indexMapping = BuildDisplayHunkToSnapshotMapping(displayHunk, currentSnapshot);
                
                System.Diagnostics.Debug.WriteLine($"[索引映射] display_hunk到当前快照的映射表:");
                System.Diagnostics.Debug.WriteLine($"[索引映射] display_hunk段落数: {GetDisplayHunkParagraphCount(displayHunk)}, 当前快照段落数: {currentSnapshot.Count}");
                foreach (var kvp in indexMapping.OrderBy(x => x.Key))
                {
                    System.Diagnostics.Debug.WriteLine($"[索引映射] display_hunk位置 {kvp.Key} -> 当前快照位置 {kvp.Value}");
                }

                // 按原始索引顺序处理操作（不排序），维护索引映射表
                foreach (var action in actions)
                {
                    try
                    {
                        // 第三步：通过start/end文本找到段落序号
                        int paragraphIndex = 0;
                        
                        // 声明range变量用于查找文本范围
                        Word.Range range = null;
                        
                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                            case "delete":
                                // 通过start和end文本找到文本范围，然后找到段落序号
                                if (!string.IsNullOrEmpty(action.StartText) && !string.IsNullOrEmpty(action.EndText))
                                {
                                    // 如果同时提供了start和end，使用FindAllText查找所有匹配
                                    var startMatches = documentProcessor.FindAllText(doc.Content, action.StartText);
                                    var endMatches = documentProcessor.FindAllText(doc.Content, action.EndText);
                                    
                                    // 检查匹配数量
                                    if (startMatches.Count == 0 || endMatches.Count == 0)
                                    {
                                        string searchText = startMatches.Count == 0 ? action.StartText : action.EndText;
                                        return $"你输入的原文本查找不到：{searchText}";
                                    }
                                    
                                    if (startMatches.Count > 1 || endMatches.Count > 1)
                                    {
                                        string searchText = startMatches.Count > 1 ? action.StartText : action.EndText;
                                        return $"通过你输入的文本查找到的段落大于1，请输入特异性更强的文本：{searchText}";
                                    }
                                    
                                    // 使用FindTextByStartEnd查找范围
                                    range = documentProcessor.FindTextByStartEnd(doc.Content, action.StartText, action.EndText);
                                    if (range != null)
                                    {
                                        paragraphIndex = FindParagraphIndexByRange(doc, range);
                                        System.Diagnostics.Debug.WriteLine($"[快照操作] {action.Type}操作：通过start='{action.StartText}'和end='{action.EndText}'找到段落序号{paragraphIndex}");
                                    }
                                    else
                                    {
                                        return $"你输入的原文本查找不到：start='{action.StartText}', end='{action.EndText}'";
                                    }
                                }
                                else if (!string.IsNullOrEmpty(action.StartText))
                                {
                                    // 如果只有start，使用FindAllText查找所有匹配
                                    var matches = documentProcessor.FindAllText(doc.Content, action.StartText);
                                    
                                    // 检查匹配数量
                                    if (matches.Count == 0)
                                    {
                                        return $"你输入的原文本查找不到：{action.StartText}";
                                    }
                                    
                                    if (matches.Count > 1)
                                    {
                                        return $"通过你输入的文本查找到的段落大于1，请输入特异性更强的文本：{action.StartText}";
                                    }
                                    
                                    // 使用FindText查找范围
                                    range = documentProcessor.FindText(doc.Content, action.StartText);
                                    if (range != null)
                                    {
                                        paragraphIndex = FindParagraphIndexByRange(doc, range);
                                        System.Diagnostics.Debug.WriteLine($"[快照操作] {action.Type}操作：通过start='{action.StartText}'找到段落序号{paragraphIndex}");
                                    }
                                    else
                                    {
                                        return $"你输入的原文本查找不到：{action.StartText}";
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] {action.Type}操作：缺少start文本");
                                    continue;
                                }
                                break;

                            case "insert":
                                // 通过pre和post文本找到插入位置，然后找到段落序号
                                if (!string.IsNullOrEmpty(action.PreText) && !string.IsNullOrEmpty(action.PostText))
                                {
                                    // 使用FindAllText查找所有匹配
                                    var preMatches = documentProcessor.FindAllText(doc.Content, action.PreText);
                                    var postMatches = documentProcessor.FindAllText(doc.Content, action.PostText);
                                    
                                    // 检查匹配数量
                                    if (preMatches.Count == 0 || postMatches.Count == 0)
                                    {
                                        string searchText = preMatches.Count == 0 ? action.PreText : action.PostText;
                                        return $"你输入的原文本查找不到：{searchText}";
                                    }
                                    
                                    if (preMatches.Count > 1 || postMatches.Count > 1)
                                    {
                                        string searchText = preMatches.Count > 1 ? action.PreText : action.PostText;
                                        return $"通过你输入的文本查找到的段落大于1，请输入特异性更强的文本：{searchText}";
                                    }
                                    
                                    range = documentProcessor.FindTextByStartEnd(doc.Content, action.PreText, action.PostText);
                                    if (range != null)
                                    {
                                        paragraphIndex = FindParagraphIndexByRange(doc, range);
                                        System.Diagnostics.Debug.WriteLine($"[快照操作] 插入操作：通过pre='{action.PreText}'和post='{action.PostText}'找到段落序号{paragraphIndex}");
                                    }
                                    else
                                    {
                                        return $"你输入的原文本查找不到：pre='{action.PreText}', post='{action.PostText}'";
                                    }
                                }
                                break;
                        }

                        if (paragraphIndex <= 0 || paragraphIndex > doc.Paragraphs.Count)
                        {
                            System.Diagnostics.Debug.WriteLine($"[快照操作] 警告：Word段落序号 {paragraphIndex} 无效，跳过操作");
                            continue;
                        }

                        // 对于replace操作，检查是否需要部分匹配
                        if (action.Type?.ToLower() == "replace" && !string.IsNullOrEmpty(action.StartText))
                        {
                            // 获取找到的段落的完整内容
                            var paragraph = doc.Paragraphs[paragraphIndex];
                            string paragraphContent = paragraph.Range.Text?.TrimEnd('\r', '\n') ?? "";
                            paragraphContent = RemoveButtonTextFromContent(doc, paragraph.Range, paragraphContent);
                            
                            // 检查段落内容是否与start文本完全匹配（去除空白字符后）
                            string normalizedParagraphContent = paragraphContent.Trim();
                            string normalizedStartText = action.StartText.Trim();
                            
                            if (!normalizedParagraphContent.Equals(normalizedStartText, StringComparison.Ordinal))
                            {
                                // 段落内容与start文本不完全匹配，尝试部分匹配
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 尝试部分匹配：搜索文本 '{action.StartText}'");
                                
                                // 检查start文本在文档中是否唯一
                                var startMatches = documentProcessor.FindAllText(doc.Content, action.StartText);
                                
                                if (startMatches.Count == 1)
                                {
                                    // start文本唯一，确认当前段落的内容包含start文本
                                    if (paragraphContent.Contains(action.StartText))
                                    {
                                        // 当前段落包含start文本，进行替换
                                        System.Diagnostics.Debug.WriteLine($"[快照操作] 找到唯一匹配的段落: {paragraphContent.Substring(0, Math.Min(100, paragraphContent.Length))}");
                                        
                                        // 获取新文本（action.Content）
                                        string newText = action.Content ?? "";
                                        
                                        // 将段落中的start文本替换为new文本
                                        string newParagraph = paragraphContent.Replace(action.StartText, newText);
                                        System.Diagnostics.Debug.WriteLine($"[快照操作] 生成新段落: {newParagraph.Substring(0, Math.Min(100, newParagraph.Length))}");
                                        
                                        // 更新action.Content为新段落，后续的replace操作会使用新段落
                                        action.Content = newParagraph;
                                        
                                        // 由于start文本是唯一的，当前找到的paragraphIndex就是正确的段落位置
                                        System.Diagnostics.Debug.WriteLine($"[快照操作] 部分匹配成功：段落内容已更新，段落位置 {paragraphIndex}");
                                    }
                                    else
                                    {
                                        // 当前段落不包含start文本，但在所有段落中搜索包含该文本的段落
                                        var matchingParagraphs = new List<(string paragraphName, string paragraphContent)>();
                                        foreach (var paraName in currentSnapshot)
                                        {
                                            string paraContent = DocumentState.GetParagraphContent(paraName);
                                            if (!string.IsNullOrEmpty(paraContent) && paraContent.Contains(action.StartText))
                                            {
                                                matchingParagraphs.Add((paraName, paraContent));
                                            }
                                        }
                                        
                                        if (matchingParagraphs.Count == 1)
                                        {
                                            // 找到唯一匹配，进行替换
                                            var (matchedParagraphName, matchedParagraphContent) = matchingParagraphs[0];
                                            System.Diagnostics.Debug.WriteLine($"[快照操作] 找到唯一匹配的段落: {matchedParagraphContent.Substring(0, Math.Min(100, matchedParagraphContent.Length))}");
                                            
                                            // 获取新文本（action.Content）
                                            string newText = action.Content ?? "";
                                            
                                            // 将段落中的start文本替换为new文本
                                            string newParagraph = matchedParagraphContent.Replace(action.StartText, newText);
                                            System.Diagnostics.Debug.WriteLine($"[快照操作] 生成新段落: {newParagraph.Substring(0, Math.Min(100, newParagraph.Length))}");
                                            
                                            // 更新action.Content为新段落，后续的replace操作会使用新段落
                                            action.Content = newParagraph;
                                            
                                            // 需要在display_hunk中找到该段落的位置
                                            int matchedDisplayHunkPosition = -1;
                                            for (int i = 0; i < displayHunk.Count; i++)
                                            {
                                                var hunkSection = displayHunk[i];
                                                for (int j = 0; j < hunkSection.lines.Count; j++)
                                                {
                                                    string line = hunkSection.lines[j];
                                                    // 去除+或-前缀
                                                    string paraName = line.StartsWith("-") || line.StartsWith("+") ? line.Substring(1) : line;
                                                    if (paraName == matchedParagraphName)
                                                    {
                                                        // 计算在display_hunk中的绝对位置
                                                        int absolutePos = 0;
                                                        for (int k = 0; k < i; k++)
                                                        {
                                                            absolutePos += displayHunk[k].lines.Count;
                                                        }
                                                        absolutePos += j + 1; // 转换为从1开始的索引
                                                        matchedDisplayHunkPosition = absolutePos;
                                                        break;
                                                    }
                                                }
                                                if (matchedDisplayHunkPosition > 0) break;
                                            }
                                            
                                            if (matchedDisplayHunkPosition > 0 && indexMapping.ContainsKey(matchedDisplayHunkPosition))
                                            {
                                                paragraphIndex = matchedDisplayHunkPosition;
                                                System.Diagnostics.Debug.WriteLine($"[快照操作] 部分匹配成功：原始段落名字 {matchedParagraphName}，新段落内容已更新，段落位置 {paragraphIndex}");
                                            }
                                            else
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[快照操作] 部分匹配失败：无法找到匹配段落在display_hunk中的位置");
                                                return "你给的原文段落不存在。请注意：本工具是以段落为级别的修改工具，你输入的文本必须是完整的段落。";
                                            }
                                        }
                                        else if (matchingParagraphs.Count > 1)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[快照操作] 部分匹配失败：找到{matchingParagraphs.Count}个匹配的段落，需要更具体的文本");
                                            return $"你给的原文文本在文档中找到多个匹配（{matchingParagraphs.Count}个），请输入更具体的文本。请注意：本工具是以段落为级别的修改工具，你输入的文本必须是完整的段落。";
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[快照操作] 部分匹配失败：未找到包含该文本的段落");
                                            return "你给的原文段落不存在。请注意：本工具是以段落为级别的修改工具，你输入的文本必须是完整的段落。";
                                        }
                                    }
                                }
                                else if (startMatches.Count > 1)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 部分匹配失败：start文本在文档中找到多个匹配（{startMatches.Count}个）");
                                    return $"你给的原文文本在文档中找到多个匹配（{startMatches.Count}个），请输入更具体的文本。请注意：本工具是以段落为级别的修改工具，你输入的文本必须是完整的段落。";
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 部分匹配失败：start文本在文档中未找到");
                                    return "你给的原文段落不存在。请注意：本工具是以段落为级别的修改工具，你输入的文本必须是完整的段落。";
                                }
                            }
                        }

                        // paragraphIndex是Word段落位置，由于display_hunk就是Word的抽象，所以它也是display_hunk位置
                        int displayHunkPosition = paragraphIndex;
                        
                        // 通过映射表找到当前快照位置
                        if (!indexMapping.ContainsKey(displayHunkPosition))
                        {
                            System.Diagnostics.Debug.WriteLine($"[索引映射] 警告：display_hunk位置 {displayHunkPosition} 不在映射表中，跳过操作");
                            continue;
                        }
                        
                        int snapshotIndex = indexMapping[displayHunkPosition];
                        System.Diagnostics.Debug.WriteLine($"[索引映射] 操作 {action.Type}：display_hunk位置 {displayHunkPosition} -> 当前快照位置 {snapshotIndex}");

                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 替换操作：display_hunk位置{displayHunkPosition} -> 当前快照位置{snapshotIndex}");
                                DocumentState.ProcessParagraphReplace(snapshotIndex, action.Content);
                                // 替换操作后，快照已更新，下次操作时会重新计算映射表
                                break;

                            case "delete":
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 删除操作：display_hunk位置{displayHunkPosition} -> 当前快照位置{snapshotIndex}");
                                DocumentState.ProcessParagraphDelete(snapshotIndex);
                                // 删除操作后，快照已更新，下次操作时会重新计算映射表
                                break;

                            case "insert":
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 插入操作：display_hunk位置{displayHunkPosition} -> 当前快照位置{snapshotIndex}, 位置{action.Position}");
                                DocumentState.ProcessParagraphInsert(snapshotIndex, action.Position, action.Content);
                                // 插入操作后，快照已更新，下次操作时会重新计算映射表
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 快照操作失败: {ex.Message}");
                    }
                }

                await Task.CompletedTask;
                return null; // 成功，返回null
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 处理快照操作失败: {ex.Message}");
                return $"处理快照操作失败: {ex.Message}";
            }
        }

        /// <summary>
        /// 建立display_hunk位置到当前快照位置的映射
        /// 使用两个指针同时遍历display_hunk和当前快照
        /// 当display_hunk指针遇到-号行时，只移动display_hunk指针，快照指针不动
        /// 当display_hunk指针遇到非-号行时，记录映射，两个指针都移动
        /// </summary>
        /// <param name="displayHunk">display_hunk</param>
        /// <param name="currentSnapshot">当前快照</param>
        /// <returns>映射表：display_hunk位置 -> 当前快照位置</returns>
        private static Dictionary<int, int> BuildDisplayHunkToSnapshotMapping(
            IReadOnlyList<(string header, IReadOnlyList<string> lines)> displayHunk,
            List<string> currentSnapshot)
        {
            var mapping = new Dictionary<int, int>();
            
            if (displayHunk == null || displayHunk.Count == 0 || currentSnapshot == null || currentSnapshot.Count == 0)
            {
                return mapping;
            }
            
            // 获取display_hunk的所有行
            var displayHunkLines = new List<string>();
            foreach (var (header, lines) in displayHunk)
            {
                foreach (var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        displayHunkLines.Add(line);
                    }
                }
            }
            
            // 两个指针同时遍历
            int displayHunkIndex = 0;  // display_hunk指针（从0开始，对应位置从1开始）
            int snapshotIndex = 0;      // 快照指针（从0开始，对应位置从1开始）
            
            while (displayHunkIndex < displayHunkLines.Count && snapshotIndex < currentSnapshot.Count)
            {
                string displayHunkLine = displayHunkLines[displayHunkIndex];
                
                // 如果display_hunk指向-号行，只移动display_hunk指针，快照指针不动
                if (displayHunkLine.StartsWith("-"))
                {
                    displayHunkIndex++;
                    continue;
                }
                
                // 如果display_hunk指向非-号行，记录映射，两个指针都移动
                // 获取段落名称（去掉+号前缀，如果有的话）
                string displayHunkParaName = displayHunkLine.StartsWith("+")
                    ? displayHunkLine.Substring(1)
                    : displayHunkLine;
                
                string snapshotParaName = currentSnapshot[snapshotIndex];
                
                // 记录映射：display_hunk位置（从1开始）-> 快照位置（从1开始）
                int displayHunkPosition = displayHunkIndex + 1;
                int snapshotPosition = snapshotIndex + 1;
                mapping[displayHunkPosition] = snapshotPosition;
                
                System.Diagnostics.Debug.WriteLine($"[映射构建] display_hunk位置{displayHunkPosition}({displayHunkParaName}) -> 快照位置{snapshotPosition}({snapshotParaName})");
                
                // 两个指针都移动
                displayHunkIndex++;
                snapshotIndex++;
            }
            
            return mapping;
        }

        /// <summary>
        /// 获取display_hunk中的段落总数
        /// </summary>
        private static int GetDisplayHunkParagraphCount(IReadOnlyList<(string header, IReadOnlyList<string> lines)> displayHunk)
        {
            int count = 0;
            if (displayHunk != null)
            {
                foreach (var (header, lines) in displayHunk)
                {
                    foreach (var line in lines)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            count++;
                        }
                    }
                }
            }
            return count;
        }

        /// <summary>
        /// 更新索引映射表：删除操作后，删除索引之后的所有索引都要-1
        /// </summary>
        /// <param name="indexMapping">索引映射表</param>
        /// <param name="deletedIndex">被删除的段落索引（从1开始）</param>
        private static void UpdateIndexMappingAfterDelete(Dictionary<int, int> indexMapping, int deletedIndex)
        {
            // 找到所有映射值 > deletedIndex 的条目，将其值-1
            var keysToUpdate = indexMapping.Where(kvp => kvp.Value > deletedIndex).Select(kvp => kvp.Key).ToList();
            foreach (var key in keysToUpdate)
            {
                int oldValue = indexMapping[key];
                indexMapping[key]--;
                System.Diagnostics.Debug.WriteLine($"[索引映射] 删除后更新：原始索引 {key} -> {oldValue} -> {indexMapping[key]}");
            }
            
            // 删除被删除的索引映射
            var keyToRemove = indexMapping.FirstOrDefault(kvp => kvp.Value == deletedIndex).Key;
            if (keyToRemove > 0)
            {
                indexMapping.Remove(keyToRemove);
                System.Diagnostics.Debug.WriteLine($"[索引映射] 移除被删除的索引映射：原始索引 {keyToRemove} -> 当前索引 {deletedIndex}");
            }
        }

        /// <summary>
        /// 更新索引映射表：插入操作后，插入点之后的所有索引都要+1
        /// </summary>
        /// <param name="indexMapping">索引映射表</param>
        /// <param name="updateStartIndex">需要更新的起始段落索引（从1开始）</param>
        private static void UpdateIndexMappingAfterInsert(Dictionary<int, int> indexMapping, int updateStartIndex)
        {
            // 找到所有映射值 >= updateStartIndex 的条目，将其值+1
            var keysToUpdate = indexMapping.Where(kvp => kvp.Value >= updateStartIndex).Select(kvp => kvp.Key).ToList();
            foreach (var key in keysToUpdate)
            {
                int oldValue = indexMapping[key];
                indexMapping[key]++;
                System.Diagnostics.Debug.WriteLine($"[索引映射] 插入后更新：原始索引 {key} -> {oldValue} -> {indexMapping[key]}");
            }
        }

        /// <summary>
        /// 根据hunk执行文档插入和删除操作
        /// 新算法：
        /// 第一步：使用LCS算法找到保持顺序一致的锚点
        /// 第二步：删除相邻锚点之间的内容（不在target对应位置的）
        /// 第三步：插入新内容（target中对应位置的内容）
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="hunk">新的全量hunk列表（只包含一个全量hunk）</param>
        /// <param name="originalFullDisplayContent">起点：从当前的display_hunk直接获取</param>
        /// <param name="targetFullDisplayContent">终点：从新的全量hunk直接获取</param>
        private static async Task ApplyHunkInsertions(Word.Document doc, List<(string header, List<string> lines)> hunk, 
            List<string> originalFullDisplayContent, List<string> targetFullDisplayContent)
        {
            try
            {
                // 构建段落名称到文档段落索引的映射（基于当前文档内容）
                var paragraphNameToIndex = new Dictionary<string, int>();
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤1: 开始构建映射表，文档段落总数: {doc.Paragraphs.Count}");
                for (int i = 1; i <= doc.Paragraphs.Count; i++)
                {
                    var para = doc.Paragraphs[i];
                    string content = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                    // 去除接受和丢弃按钮文本后再进行匹配
                    content = RemoveButtonTextFromContent(doc, para.Range, content);
                    // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                    string paraName = DocumentState.FindParagraphNameByContent(content);
                    if (paraName != null && !paragraphNameToIndex.ContainsKey(paraName))
                    {
                        paragraphNameToIndex[paraName] = i;
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤1完成: 映射表构建完成，共 {paragraphNameToIndex.Count} 个映射");
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] originalFullDisplayContent: {string.Join(",", originalFullDisplayContent)}");
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] targetFullDisplayContent: {string.Join(",", targetFullDisplayContent)}");

                // 第一步：使用LCS算法找到保持顺序一致的唯一锚点序列
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤2: 开始使用LCS算法找到唯一锚点序列");
                var anchorSequences = FindOrderedAnchorPoints(originalFullDisplayContent, targetFullDisplayContent);
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤2完成: 找到 {anchorSequences.Count} 个唯一锚点序列");

                // 第二步：删除相邻锚点之间的内容
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤3: 开始删除相邻锚点之间的内容");
                var deletions = new List<int>();

                // 遍历相邻的锚点序列对
                for (int i = 0; i < anchorSequences.Count - 1; i++)
                {
                    var startSequence = anchorSequences[i];
                    var endSequence = anchorSequences[i + 1];

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 处理锚点序列区间 [{string.Join(",", startSequence.sequence)}({startSequence.targetStartIndex}-{startSequence.targetEndIndex}) - {string.Join(",", endSequence.sequence)}({endSequence.targetStartIndex}-{endSequence.targetEndIndex})]");

                    // 使用锚点在原始文档中的位置来确定文档中的位置
                    // 注意：originalEndIndex 和 originalStartIndex 是数组索引（从0开始）
                    // 文档段落索引是从1开始的
                    int startDocIndex = startSequence.originalEndIndex + 2; // 锚点序列结束后的第一个段落（数组索引+1对应文档索引+1）
                    int endDocIndex = endSequence.originalStartIndex + 1; // 下一个锚点序列开始的段落（数组索引+1对应文档索引）

                    // 收集target中这个区间应该有的段落
                    var targetSegment = new HashSet<string>();
                    for (int j = startSequence.targetEndIndex + 1; j < endSequence.targetStartIndex; j++)
                    {
                        if (j < targetFullDisplayContent.Count)
                        {
                            targetSegment.Add(targetFullDisplayContent[j]);
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] target区间 [{startSequence.targetEndIndex + 1} - {endSequence.targetStartIndex - 1}] 应包含: {string.Join(",", targetSegment)}");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 文档区间检查: 从段落{startDocIndex}到{endDocIndex - 1}");

                    // 检查文档中这个区间内的段落
                    // 如果target区间为空（两个锚点在target中相邻），则删除original中对应区间的所有段落
                    // 如果target区间不为空，则删除不在targetSegment中的段落
                    for (int docIdx = startDocIndex; docIdx < endDocIndex; docIdx++)
                    {
                        if (docIdx < 1 || docIdx > doc.Paragraphs.Count)
                            continue;

                        var para = doc.Paragraphs[docIdx];
                        string docContent = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                        docContent = RemoveButtonTextFromContent(doc, para.Range, docContent);

                        // 找到这个段落对应的段落名称（使用哈希码验证）
                        string paraName = DocumentState.FindParagraphNameByContent(docContent);

                        if (paraName != null)
                        {
                            // 如果target区间为空，或者这个段落不在target的对应区间中，需要删除
                            if (targetSegment.Count == 0 || !targetSegment.Contains(paraName))
                            {
                                deletions.Add(docIdx);
                                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 标记删除段落 {paraName} (文档索引: {docIdx})，{(targetSegment.Count == 0 ? "target区间为空" : "不在target区间中")}");
                            }
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤3完成: 共标记 {deletions.Count} 个段落需要删除");
                
                // 处理最后一个锚点序列之后的内容
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤4: 处理最后一个锚点序列之后的内容");
                if (anchorSequences.Count > 0)
                {
                    var lastSequence = anchorSequences[anchorSequences.Count - 1];

                    // 使用锚点在原始文档中的结束位置
                    int lastDocIndex = lastSequence.originalEndIndex;

                    var targetAfterLastAnchor = new HashSet<string>();
                    for (int j = lastSequence.targetEndIndex + 1; j < targetFullDisplayContent.Count; j++)
                    {
                        targetAfterLastAnchor.Add(targetFullDisplayContent[j]);
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点序列 {string.Join(",", lastSequence.sequence)} (original结束位置: {lastDocIndex}) 之后应包含: {string.Join(",", targetAfterLastAnchor)}");

                    // 只有当target中确实指定了最后一个锚点之后的内容时，才执行删除逻辑
                    if (targetAfterLastAnchor.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 检查文档从段落{lastDocIndex + 1}到文档末尾");

                        // 删除lastDocIndex之后不在target中的段落
                        for (int docIdx = lastDocIndex + 1; docIdx <= doc.Paragraphs.Count; docIdx++)
                        {
                            var para = doc.Paragraphs[docIdx];
                            string docContent = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                            docContent = RemoveButtonTextFromContent(doc, para.Range, docContent);

                            // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                            string paraName = DocumentState.FindParagraphNameByContent(docContent);

                            if (paraName != null && !targetAfterLastAnchor.Contains(paraName))
                            {
                                deletions.Add(docIdx);
                                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 标记删除最后一个锚点序列后的段落 {paraName} (文档索引: {docIdx})");
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点序列已是文档末尾，无需删除操作");
                    }
                }

                // 执行删除操作（从后往前）
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤5: 执行删除操作，共找到 {deletions.Count} 个需要删除的段落");
                deletions.Sort((a, b) => b.CompareTo(a)); // 降序排序
                foreach (var docIndex in deletions)
                {
                    if (docIndex >= 1 && docIndex <= doc.Paragraphs.Count)
                    {
                        var paragraph = doc.Paragraphs[docIndex];
                        paragraph.Range.Delete();
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 已删除文档段落{docIndex}");

                        // 更新映射表：删除后，所有大于docIndex的索引都要-1
                        var keysToUpdate = paragraphNameToIndex.Where(kvp => kvp.Value > docIndex).Select(kvp => kvp.Key).ToList();
                        foreach (var key in keysToUpdate)
                        {
                            paragraphNameToIndex[key]--;
                        }
                        // 移除被删除的映射
                        var keyToRemove = paragraphNameToIndex.FirstOrDefault(kvp => kvp.Value == docIndex).Key;
                        if (!string.IsNullOrEmpty(keyToRemove))
                        {
                            paragraphNameToIndex.Remove(keyToRemove);
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤5完成: 删除操作完成，当前文档段落总数: {doc.Paragraphs.Count}");

                // 第三步：插入新内容，按照target的顺序在相邻锚点之间插入缺失的段落
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤6: 重新构建映射表（删除后）");
                // 重新构建映射表（删除后）
                paragraphNameToIndex.Clear();
                for (int i = 1; i <= doc.Paragraphs.Count; i++)
                {
                    var para = doc.Paragraphs[i];
                    string content = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                    content = RemoveButtonTextFromContent(doc, para.Range, content);
                    // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                    string paraName = DocumentState.FindParagraphNameByContent(content);
                    if (paraName != null && !paragraphNameToIndex.ContainsKey(paraName))
                    {
                        paragraphNameToIndex[paraName] = i;
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤6完成: 重新构建映射表完成，当前映射数量: {paragraphNameToIndex.Count}");
                
                // 按照target的顺序，在相邻锚点序列之间插入缺失的段落
                // 从后往前遍历锚点区间，避免前面插入影响后面位置
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤7: 按照target的顺序插入缺失的段落（从后往前处理）");
                for (int intervalIndex = anchorSequences.Count - 2; intervalIndex >= 0; intervalIndex--)
                {
                    var startSequence = anchorSequences[intervalIndex];
                    var endSequence = anchorSequences[intervalIndex + 1];

                    // 使用paragraphNameToIndex映射表找到锚点序列在删除后文档中的实际位置
                    // 在startSequence结束之后开始插入
                    int currentInsertPosition = -1;
                    if (startSequence.sequence.Count > 0)
                    {
                        // 找到锚点序列的最后一个段落在当前文档中的位置
                        var lastAnchorPara = startSequence.sequence.Last();
                        if (paragraphNameToIndex.ContainsKey(lastAnchorPara))
                        {
                            currentInsertPosition = paragraphNameToIndex[lastAnchorPara];
                        }
                    }

                    // 收集target中这个区间应该有的段落（按顺序）
                    var targetSegment = new List<string>();
                    for (int j = startSequence.targetEndIndex + 1; j < endSequence.targetStartIndex; j++)
                    {
                        if (j < targetFullDisplayContent.Count)
                        {
                            targetSegment.Add(targetFullDisplayContent[j]);
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 插入区间 [{string.Join(",", startSequence.sequence)} - {string.Join(",", endSequence.sequence)}] 应包含: {string.Join(",", targetSegment)}");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 在锚点序列结束位置{currentInsertPosition}之后插入");

                    // 找出哪些段落已经在文档中存在（在锚点序列之间的区间内）
                    var existingInDoc = new HashSet<string>();
                    if (currentInsertPosition >= 0)
                    {
                        // 找到下一个锚点序列的开始位置
                        int nextAnchorPosition = -1;
                        if (endSequence.sequence.Count > 0)
                        {
                            var firstNextAnchorPara = endSequence.sequence.First();
                            if (paragraphNameToIndex.ContainsKey(firstNextAnchorPara))
                            {
                                nextAnchorPosition = paragraphNameToIndex[firstNextAnchorPara];
                            }
                        }

                        // 检查当前锚点序列结束位置到下一个锚点序列开始位置之间的段落
                        int startCheckIdx = currentInsertPosition + 1;
                        int endCheckIdx = (nextAnchorPosition >= 0) ? nextAnchorPosition : doc.Paragraphs.Count + 1;
                        for (int docIdx = startCheckIdx; docIdx < endCheckIdx; docIdx++)
                        {
                            if (docIdx < 1 || docIdx > doc.Paragraphs.Count)
                                continue;

                            var para = doc.Paragraphs[docIdx];
                            string docContent = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                            docContent = RemoveButtonTextFromContent(doc, para.Range, docContent);

                            // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                            string paraName = DocumentState.FindParagraphNameByContent(docContent);
                            if (paraName != null && targetSegment.Contains(paraName))
                            {
                                existingInDoc.Add(paraName);
                            }
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 已存在的段落: {string.Join(",", existingInDoc)}");

                    // 区间内从前往后插入，保持正确的顺序
                    // 由于区间之间是从后往前处理的，所以前面区间的插入不会影响当前区间的位置
                    foreach (var paraName in targetSegment)
                    {
                        if (!existingInDoc.Contains(paraName))
                        {
                            // 需要插入这个段落
                            string content = DocumentState.GetParagraphContent(paraName);
                            if (!string.IsNullOrEmpty(content))
                            {
                                if (currentInsertPosition >= 1 && currentInsertPosition <= doc.Paragraphs.Count)
                                {
                                    // 在锚点序列结束位置之后插入新段落
                                    // currentInsertPosition是锚点序列最后一个段落的文档索引（从1开始）
                                    // 我们在当前段落的结束位置之后插入
                                    Word.Range insertRange;
                                    if (currentInsertPosition <= doc.Paragraphs.Count)
                                    {
                                        // 在锚点序列结束的段落结束位置之后插入
                                        var anchorParagraph = doc.Paragraphs[currentInsertPosition];
                                        insertRange = anchorParagraph.Range;
                                        insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 在段落{currentInsertPosition}结束处插入: {paraName}");
                                    }
                                    else
                                    {
                                        // 在文档末尾插入
                                        insertRange = doc.Range(doc.Content.End - 1, doc.Content.End);
                                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 在文档末尾插入: {paraName}");
                                    }
                                    insertRange.Text = content + "\r";

                                    // 更新映射表 - 新插入的段落位于currentInsertPosition + 1的位置
                                    paragraphNameToIndex[paraName] = currentInsertPosition + 1;
                                    // 更新后续索引
                                    var keysToUpdate = paragraphNameToIndex.Where(kvp => kvp.Value > currentInsertPosition + 1 && kvp.Key != paraName).Select(kvp => kvp.Key).ToList();
                                    foreach (var key in keysToUpdate)
                                    {
                                        paragraphNameToIndex[key]++;
                                    }

                                    // 更新插入位置，下次插入就在这个新插入的段落之后
                                    currentInsertPosition++;
                                }
                            }
                        }
                        else
                        {
                            // 段落已存在，更新插入位置
                            if (paragraphNameToIndex.ContainsKey(paraName))
                            {
                                currentInsertPosition = paragraphNameToIndex[paraName];
                            }
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤7完成: 锚点间插入完成，当前文档段落总数: {doc.Paragraphs.Count}");
                
                // 处理最后一个锚点序列之后的内容
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤8: 处理最后一个锚点序列之后的内容");
                if (anchorSequences.Count > 0)
                {
                    var lastSequence = anchorSequences[anchorSequences.Count - 1];

                    // 使用paragraphNameToIndex映射表找到最后一个锚点序列在删除后文档中的实际位置
                    int currentInsertPosition = -1;
                    if (lastSequence.sequence.Count > 0)
                    {
                        // 找到锚点序列的最后一个段落在当前文档中的位置
                        var lastAnchorPara = lastSequence.sequence.Last();
                        if (paragraphNameToIndex.ContainsKey(lastAnchorPara))
                        {
                            currentInsertPosition = paragraphNameToIndex[lastAnchorPara];
                        }
                    }

                    // 收集target中最后一个锚点序列之后应该有的段落
                    var targetAfterLastAnchor = new List<string>();
                    for (int j = lastSequence.targetEndIndex + 1; j < targetFullDisplayContent.Count; j++)
                    {
                        targetAfterLastAnchor.Add(targetFullDisplayContent[j]);
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点序列 {string.Join(",", lastSequence.sequence)} (文档位置: {currentInsertPosition}) 之后应插入: {string.Join(",", targetAfterLastAnchor)}");

                    // 找出哪些段落已经在文档中存在（在最后一个锚点之后）
                    var existingInDoc = new HashSet<string>();
                    if (currentInsertPosition >= 1)
                    {
                        for (int docIdx = currentInsertPosition + 1; docIdx <= doc.Paragraphs.Count; docIdx++)
                        {
                            var para = doc.Paragraphs[docIdx];
                            string docContent = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                            docContent = RemoveButtonTextFromContent(doc, para.Range, docContent);

                            // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                            string paraName = DocumentState.FindParagraphNameByContent(docContent);
                            if (paraName != null && targetAfterLastAnchor.Contains(paraName))
                            {
                                existingInDoc.Add(paraName);
                            }
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点序列之后已存在的段落: {string.Join(",", existingInDoc)}");

                    // 从前往后插入，保持正确的顺序，每次插入后更新位置
                    foreach (var paraName in targetAfterLastAnchor)
                    {
                        if (!existingInDoc.Contains(paraName))
                        {
                            string content = DocumentState.GetParagraphContent(paraName);
                            if (!string.IsNullOrEmpty(content))
                            {
                                if (currentInsertPosition >= 1 && currentInsertPosition <= doc.Paragraphs.Count)
                                {
                                    // 在最后一个锚点序列结束位置之后插入新段落
                                    Word.Range insertRange;
                                    if (currentInsertPosition <= doc.Paragraphs.Count)
                                    {
                                        // 在锚点序列结束的段落结束位置之后插入
                                        var anchorParagraph = doc.Paragraphs[currentInsertPosition];
                                        insertRange = anchorParagraph.Range;
                                        insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 在段落{currentInsertPosition}结束处插入: {paraName}");
                                    }
                                    else
                                    {
                                        // 在文档末尾插入
                                        insertRange = doc.Range(doc.Content.End - 1, doc.Content.End);
                                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 在文档末尾插入: {paraName}");
                                    }
                                    insertRange.Text = content + "\r";

                                    // 更新映射表
                                    paragraphNameToIndex[paraName] = currentInsertPosition + 1;
                                    var keysToUpdate = paragraphNameToIndex.Where(kvp => kvp.Value > currentInsertPosition + 1 && kvp.Key != paraName).Select(kvp => kvp.Key).ToList();
                                    foreach (var key in keysToUpdate)
                                    {
                                        paragraphNameToIndex[key]++;
                                    }

                                    // 更新插入位置，下次插入就在这个新插入的段落之后
                                    currentInsertPosition++;
                                }
                            }
                        }
                        else
                        {
                            // 段落已存在，更新插入位置
                            if (paragraphNameToIndex.ContainsKey(paraName))
                            {
                                currentInsertPosition = paragraphNameToIndex[paraName];
                            }
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤8完成: 所有插入操作完成，最终文档段落总数: {doc.Paragraphs.Count}");

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用Hunk插入操作失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用改进的LCS算法找到保持顺序一致的唯一锚点序列
        /// 返回按target顺序排列的锚点列表，每个锚点包含原始和目标中的位置及序列
        /// </summary>
        /// <param name="original">原始段落列表</param>
        /// <param name="target">目标段落列表</param>
        /// <returns>锚点列表，按target顺序排列</returns>
        private static List<(int originalStartIndex, int originalEndIndex, int targetStartIndex, int targetEndIndex, List<string> sequence)> FindOrderedAnchorPoints(List<string> original, List<string> target)
        {
            var result = new List<(int originalStartIndex, int originalEndIndex, int targetStartIndex, int targetEndIndex, List<string> sequence)>();

            if (original.Count == 0 || target.Count == 0)
                return result;

            // 使用动态规划计算LCS
            int m = original.Count;
            int n = target.Count;
            int[,] dp = new int[m + 1, n + 1];

            // 填充DP表
            for (int i = 1; i <= m; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    if (original[i - 1] == target[j - 1])
                    {
                        dp[i, j] = dp[i - 1, j - 1] + 1;
                    }
                    else
                    {
                        dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                    }
                }
            }

            // 回溯找到LCS序列
            int x = m, y = n;
            var lcsIndices = new List<(int originalIndex, int targetIndex)>();

            while (x > 0 && y > 0)
            {
                if (original[x - 1] == target[y - 1])
                {
                    // 找到公共元素
                    lcsIndices.Add((x - 1, y - 1));
                    x--;
                    y--;
                }
                else if (dp[x - 1, y] > dp[x, y - 1])
                {
                    x--;
                }
                else
                {
                    y--;
                }
            }

            // 反转，使其按target顺序排列
            lcsIndices.Reverse();

            System.Diagnostics.Debug.WriteLine($"[FindOrderedAnchorPoints] LCS找到 {lcsIndices.Count} 个匹配点");

            // 将连续的LCS点合并为唯一锚点序列
            var anchorSequences = new List<(int originalStartIndex, int originalEndIndex, int targetStartIndex, int targetEndIndex, List<string> sequence)>();

            int sequenceIndex = 0;
            while (sequenceIndex < lcsIndices.Count)
            {
                int startOriginalIdx = lcsIndices[sequenceIndex].originalIndex;
                int startTargetIdx = lcsIndices[sequenceIndex].targetIndex;
                int currentLength = 1;

                // 尝试扩展序列长度，直到找到唯一的序列
                while (sequenceIndex + currentLength <= lcsIndices.Count)
                {
                    // 检查当前长度的序列是否唯一
                    var currentSequence = new List<string>();
                    bool isValid = true;

                    // 构建当前候选序列
                    for (int k = 0; k < currentLength; k++)
                    {
                        if (sequenceIndex + k >= lcsIndices.Count)
                        {
                            isValid = false;
                            break;
                        }
                        int targetIdx = lcsIndices[sequenceIndex + k].targetIndex;
                        int originalIdx = lcsIndices[sequenceIndex + k].originalIndex;
                        if (targetIdx >= target.Count || originalIdx >= original.Count)
                        {
                            isValid = false;
                            break;
                        }
                        currentSequence.Add(target[targetIdx]);
                    }

                    if (!isValid || currentSequence.Count == 0)
                        break;

                    // 检查这个序列在original和target中是否都唯一出现
                    if (IsUniqueSequence(original, currentSequence) && IsUniqueSequence(target, currentSequence))
                    {
                        // 找到了唯一的序列，添加到结果中
                        int endOriginalIdx = lcsIndices[sequenceIndex + currentLength - 1].originalIndex;
                        int endTargetIdx = lcsIndices[sequenceIndex + currentLength - 1].targetIndex;
                        anchorSequences.Add((startOriginalIdx, endOriginalIdx, startTargetIdx, endTargetIdx, new List<string>(currentSequence)));

                        System.Diagnostics.Debug.WriteLine($"[FindOrderedAnchorPoints] 找到唯一锚点序列: {string.Join(",", currentSequence)} (original: {startOriginalIdx}-{endOriginalIdx}, target: {startTargetIdx}-{endTargetIdx})");

                        // 跳过这个序列中包含的所有点
                        sequenceIndex += currentLength;
                        break;
                    }
                    else
                    {
                        // 序列不唯一，尝试更长的序列
                        currentLength++;
                        if (sequenceIndex + currentLength > lcsIndices.Count)
                        {
                            // 如果无法找到唯一序列，使用当前的最长序列
                            int endOriginalIdx = lcsIndices[sequenceIndex + currentLength - 2].originalIndex;
                            int endTargetIdx = lcsIndices[sequenceIndex + currentLength - 2].targetIndex;
                            anchorSequences.Add((startOriginalIdx, endOriginalIdx, startTargetIdx, endTargetIdx, new List<string>(currentSequence.Take(currentLength - 1))));

                            System.Diagnostics.Debug.WriteLine($"[FindOrderedAnchorPoints] 使用最长序列作为锚点: {string.Join(",", currentSequence.Take(currentLength - 1))} (original: {startOriginalIdx}-{endOriginalIdx}, target: {startTargetIdx}-{endTargetIdx})");

                            sequenceIndex += currentLength - 1;
                            break;
                        }
                    }
                }

                if (sequenceIndex < lcsIndices.Count && lcsIndices[sequenceIndex].targetIndex == startTargetIdx)
                {
                    // 如果没有找到合适的序列，跳过这个点
                    sequenceIndex++;
                }
            }

            System.Diagnostics.Debug.WriteLine($"[FindOrderedAnchorPoints] 最终找到 {anchorSequences.Count} 个唯一锚点序列");

            return anchorSequences;
        }

        /// <summary>
        /// 检查序列在列表中是否唯一出现
        /// </summary>
        private static bool IsUniqueSequence(List<string> list, List<string> sequence)
        {
            if (sequence.Count == 0)
                return false;

            int firstMatch = FindSequenceIndex(list, sequence, 0);
            if (firstMatch == -1)
                return false;

            // 检查是否还有其他匹配
            int secondMatch = FindSequenceIndex(list, sequence, firstMatch + sequence.Count);
            return secondMatch == -1;
        }

        /// <summary>
        /// 在列表中查找序列的起始索引
        /// </summary>
        private static int FindSequenceIndex(List<string> list, List<string> sequence, int startIndex = 0)
        {
            if (sequence.Count == 0 || list.Count < sequence.Count)
                return -1;

            for (int i = startIndex; i <= list.Count - sequence.Count; i++)
            {
                bool match = true;
                for (int j = 0; j < sequence.Count; j++)
                {
                    if (list[i + j] != sequence[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// 在文档中查找序列的起始段落索引
        /// </summary>
        private static int FindSequenceStartInDocument(Word.Document doc, List<string> sequence, Dictionary<string, int> paragraphNameToIndex)
        {
            if (sequence.Count == 0)
                return -1;

            // 找到序列中第一个段落在文档中的位置
            if (paragraphNameToIndex.ContainsKey(sequence[0]))
            {
                int startIndex = paragraphNameToIndex[sequence[0]];
                // 验证后续段落是否匹配整个序列
                bool match = true;
                for (int i = 0; i < sequence.Count; i++)
                {
                    if (startIndex + i > doc.Paragraphs.Count)
                    {
                        match = false;
                        break;
                    }
                    var para = doc.Paragraphs[startIndex + i];
                    string content = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                    content = RemoveButtonTextFromContent(doc, para.Range, content);
                    string paraName = DocumentState.FindParagraphNameByContent(content);
                    if (paraName != sequence[i])
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                    return startIndex;
            }
            return -1;
        }

        /// <summary>
        /// 在文档中查找序列的结束段落索引
        /// </summary>
        private static int FindSequenceEndInDocument(Word.Document doc, List<string> sequence, Dictionary<string, int> paragraphNameToIndex)
        {
            if (sequence.Count == 0)
                return -1;

            // 找到序列中最后一个段落在文档中的位置
            if (paragraphNameToIndex.ContainsKey(sequence[sequence.Count - 1]))
            {
                int endIndex = paragraphNameToIndex[sequence[sequence.Count - 1]];
                int startIndex = endIndex - sequence.Count + 1;

                // 验证从startIndex开始的段落是否匹配整个序列
                bool match = true;
                for (int i = 0; i < sequence.Count; i++)
                {
                    if (startIndex + i < 1 || startIndex + i > doc.Paragraphs.Count)
                    {
                        match = false;
                        break;
                    }
                    var para = doc.Paragraphs[startIndex + i];
                    string content = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                    content = RemoveButtonTextFromContent(doc, para.Range, content);
                    string paraName = DocumentState.FindParagraphNameByContent(content);
                    if (paraName != sequence[i])
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                    return endIndex;
            }
            return -1;
        }

        /// <summary>
        /// 操作块类
        /// </summary>
        private class OperationBlock
        {
            public string Type { get; set; } // "replace", "delete", "insert"
            public List<string> Lines { get; set; } // hunk行
            public int StartIndex { get; set; } // 在hunkLines中的起始索引
            public int EndIndex { get; set; } // 在hunkLines中的结束索引
        }

        /// <summary>
        /// 根据全量hunk修改格式并添加按钮
        /// 全量hunk包含所有段落（未改动的和改动的），操作块识别直接在全量hunk上进行
        /// </summary>
        private static void ApplyHunkFormatting(Word.Document doc, List<string> hunkLines, List<string> targetFullDisplayContent)
        {
            try
            {
                // 第4-1步：识别操作块（在全量hunk上识别连续的改动行）
                System.Diagnostics.Debug.WriteLine("[第四步-1] 识别操作块（使用全量hunk）");
                var operationBlocks = IdentifyOperationBlocks(hunkLines);
                System.Diagnostics.Debug.WriteLine($"[第四步-1完成] 识别到 {operationBlocks.Count} 个操作块");

                // 构建段落名称到文档段落索引的映射（基于当前文档内容）
                var paragraphNameToIndex = new Dictionary<string, int>();
                for (int i = 1; i <= doc.Paragraphs.Count; i++)
                {
                    var para = doc.Paragraphs[i];
                    string content = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                    // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                    string paraName = DocumentState.FindParagraphNameByContent(content);
                    if (paraName != null && !paragraphNameToIndex.ContainsKey(paraName))
                    {
                        paragraphNameToIndex[paraName] = i;
                    }
                }

                // 第4-2步：应用格式和添加按钮
                System.Diagnostics.Debug.WriteLine("[第四步-2] 应用格式和添加按钮");
                foreach (var block in operationBlocks)
                {
                    ApplyOperationBlockFormatting(doc, block, paragraphNameToIndex, targetFullDisplayContent);
                }
                System.Diagnostics.Debug.WriteLine("[第四步-2完成]");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用Hunk格式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 识别操作块（在全量hunk上识别）
        /// 全量hunk包含所有段落，未改动的段落没有+或-前缀
        /// 操作块识别规则：
        /// 1. 连续的改动行（+或-开头的行）组成一个操作块
        /// 2. 当遇到未改动的段落（没有+或-的行）时，如果之前有改动块，就结束当前操作块
        /// 3. 当遇到新的改动行时，如果之前已经结束了一个操作块，就开始新的操作块
        /// </summary>
        private static List<OperationBlock> IdentifyOperationBlocks(List<string> hunkLines)
        {
            var blocks = new List<OperationBlock>();
            
            int startIndex = -1; // 当前操作块的开始索引

            for (int i = 0; i < hunkLines.Count; i++)
            {
                string line = hunkLines[i];
                bool isChangeLine = line.StartsWith("-") || line.StartsWith("+");
                bool isSpaceLine = !isChangeLine && !string.IsNullOrWhiteSpace(line); // 空格行：既没有+也没有-的非空行
                
                if (isChangeLine)
                {
                    // 当前行是改动行
                    if (startIndex == -1)
                    {
                        // 开始新的操作块
                        startIndex = i;
                    }
                    // 如果startIndex != -1，继续当前操作块
                }
                else if (isSpaceLine)
                {
                    // 当前行是空格行（上下文行）
                    if (startIndex >= 0)
                    {
                        // 结束当前操作块
                        var blockLines = hunkLines.GetRange(startIndex, i - startIndex);
                        var blockType = IdentifyBlockType(blockLines);
                        blocks.Add(new OperationBlock
                        {
                            Type = blockType,
                            Lines = blockLines,
                            StartIndex = startIndex,
                            EndIndex = i - 1
                        });
                        startIndex = -1;
                    }
                }
                // 如果是空行（string.IsNullOrWhiteSpace），忽略它，不影响操作块的识别
            }

            // 处理最后一个操作块（如果hunk以改动行结尾）
            if (startIndex >= 0)
            {
                var blockLines = hunkLines.GetRange(startIndex, hunkLines.Count - startIndex);
                var blockType = IdentifyBlockType(blockLines);
                blocks.Add(new OperationBlock
                {
                    Type = blockType,
                    Lines = blockLines,
                    StartIndex = startIndex,
                    EndIndex = hunkLines.Count - 1
                });
            }

            return blocks;
        }

        /// <summary>
        /// 识别操作块类型
        /// </summary>
        private static string IdentifyBlockType(List<string> blockLines)
        {
            bool hasDelete = false;
            bool hasInsert = false;

            foreach (var line in blockLines)
            {
                if (line.StartsWith("-"))
                {
                    hasDelete = true;
                }
                else if (line.StartsWith("+"))
                {
                    hasInsert = true;
                }
            }

            if (hasDelete && hasInsert)
            {
                return "replace"; // 替换操作块
            }
            else if (hasDelete)
            {
                return "delete"; // 删除操作块
            }
            else if (hasInsert)
            {
                return "insert"; // 插入操作块
            }

            return "unknown";
        }

        /// <summary>
        /// 获取段落名称对应的文档段落索引（优先使用paragraphNameToIndex，如果找不到则使用targetFullDisplayContent）
        /// </summary>
        private static int? GetParagraphDocIndex(string paragraphName, Dictionary<string, int> paragraphNameToIndex, List<string> targetFullDisplayContent)
        {
            // 优先使用paragraphNameToIndex
            if (paragraphNameToIndex.ContainsKey(paragraphName))
            {
                return paragraphNameToIndex[paragraphName];
            }
            
            // 如果找不到，使用targetFullDisplayContent中的序号
            if (targetFullDisplayContent != null)
            {
                int index = targetFullDisplayContent.IndexOf(paragraphName);
                if (index >= 0)
                {
                    // targetFullDisplayContent中的索引是从0开始的，文档段落索引是从1开始的
                    return index + 1;
                }
            }
            
            return null;
        }

        /// <summary>
        /// 应用操作块格式和添加按钮
        /// </summary>
        private static void ApplyOperationBlockFormatting(Word.Document doc, OperationBlock block, Dictionary<string, int> paragraphNameToIndex, List<string> targetFullDisplayContent)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[操作块] 类型: {block.Type}, 行数: {block.Lines.Count}");

                Word.Range operationRange = null;
                Word.Range originalRange = null;
                Word.Range newRange = null;

                // 收集操作块中涉及的段落范围
                var deleteRanges = new List<Word.Range>();
                var insertRanges = new List<Word.Range>();

                // 收集操作块中涉及的段落范围
                foreach (var line in block.Lines)
                {
                    if (line.StartsWith("-"))
                    {
                        // 负号的段落加上删除线
                        string paragraphName = line.Substring(1);
                        int? docIndex = GetParagraphDocIndex(paragraphName, paragraphNameToIndex, targetFullDisplayContent);
                        if (docIndex.HasValue && docIndex.Value >= 1 && docIndex.Value <= doc.Paragraphs.Count)
                        {
                            var paragraph = doc.Paragraphs[docIndex.Value];
                            deleteRanges.Add(paragraph.Range);
                        }
                    }
                    else if (line.StartsWith("+"))
                    {
                        // 正号的段落加上黄色背景
                        string paragraphName = line.Substring(1);
                        int? docIndex = GetParagraphDocIndex(paragraphName, paragraphNameToIndex, targetFullDisplayContent);
                        if (docIndex.HasValue && docIndex.Value >= 1 && docIndex.Value <= doc.Paragraphs.Count)
                        {
                            var paragraph = doc.Paragraphs[docIndex.Value];
                            insertRanges.Add(paragraph.Range);
                        }
                    }
                }

                // 确定操作范围（用于添加按钮）
                if (insertRanges.Count > 0)
                {
                    // 如果有插入范围，按钮放在最后一个插入段落的末尾
                    operationRange = insertRanges[insertRanges.Count - 1];
                    newRange = operationRange;
                }
                else if (deleteRanges.Count > 0)
                {
                    // 如果只有删除范围，按钮放在最后一个删除段落的末尾
                    operationRange = deleteRanges[deleteRanges.Count - 1];
                    originalRange = operationRange;
                }

                // 合并多个范围（如果有）
                if (deleteRanges.Count > 1)
                {
                    originalRange = doc.Range(deleteRanges[0].Start, deleteRanges[deleteRanges.Count - 1].End);
                }
                else if (deleteRanges.Count == 1)
                {
                    originalRange = deleteRanges[0];
                }

                if (insertRanges.Count > 1)
                {
                    newRange = doc.Range(insertRanges[0].Start, insertRanges[insertRanges.Count - 1].End);
                }
                else if (insertRanges.Count == 1)
                {
                    newRange = insertRanges[0];
                }

                // 先插入按钮（在操作块最后）
                int buttonStartPos = -1;
                if (operationRange != null)
                {
                    try
                    {
                        // 创建一个不包含段落标记的范围（在段落内容末尾插入按钮）
                        // 检查范围末尾是否有段落标记
                        string rangeText = operationRange.Text ?? "";
                        Word.Range buttonRange;
                        
                        if (rangeText.EndsWith("\r") || rangeText.EndsWith("\n"))
                        {
                            // 创建一个新的范围，排除最后一个字符（段落标记）
                            // 这样按钮会在段落内容末尾插入，然后段落标记在按钮之后
                            int endPos = operationRange.End - 1;
                            buttonRange = doc.Range(operationRange.Start, endPos);
                        }
                        else
                        {
                            // 如果没有段落标记，直接使用原范围
                            buttonRange = operationRange;
                        }
                        
                        // 记录按钮插入位置
                        buttonStartPos = buttonRange.End;
                        
                        // 创建DocumentProcessor实例来添加按钮
                        var documentProcessor = new DocumentProcessor(doc.Application);
                        documentProcessor.InsertInlineButtons(buttonRange, block.Type, originalRange, newRange);
                        System.Diagnostics.Debug.WriteLine($"[操作块按钮] {block.Type}操作块按钮添加完成，按钮位置: {buttonRange.Start}-{buttonRange.End}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 添加操作块按钮失败: {ex.Message}");
                    }
                }

                // 然后应用格式（确保格式只应用到按钮之前的内容）
                foreach (var line in block.Lines)
                {
                    if (line.StartsWith("-"))
                    {
                        // 负号的段落加上删除线
                        string paragraphName = line.Substring(1);
                        int? docIndex = GetParagraphDocIndex(paragraphName, paragraphNameToIndex, targetFullDisplayContent);
                        if (docIndex.HasValue && docIndex.Value >= 1 && docIndex.Value <= doc.Paragraphs.Count)
                        {
                            var paragraph = doc.Paragraphs[docIndex.Value];
                            
                            // 创建格式范围，确保不超过按钮位置
                            Word.Range formatRange = paragraph.Range;
                            if (buttonStartPos > 0 && formatRange.End > buttonStartPos)
                            {
                                // 如果格式范围延伸到按钮之后，限制到按钮之前
                                formatRange = doc.Range(formatRange.Start, buttonStartPos);
                            }
                            
                            ApplyDeleteFormatting(formatRange);
                            System.Diagnostics.Debug.WriteLine($"[操作块格式] 段落{paragraphName}(文档索引{docIndex.Value})添加删除线");
                        }
                    }
                    else if (line.StartsWith("+"))
                    {
                        // 正号的段落加上黄色背景
                        string paragraphName = line.Substring(1);
                        int? docIndex = GetParagraphDocIndex(paragraphName, paragraphNameToIndex, targetFullDisplayContent);
                        if (docIndex.HasValue && docIndex.Value >= 1 && docIndex.Value <= doc.Paragraphs.Count)
                        {
                            var paragraph = doc.Paragraphs[docIndex.Value];
                            
                            // 创建格式范围，确保不超过按钮位置
                            Word.Range formatRange = paragraph.Range;
                            if (buttonStartPos > 0 && formatRange.End > buttonStartPos)
                            {
                                // 如果格式范围延伸到按钮之后，限制到按钮之前
                                formatRange = doc.Range(formatRange.Start, buttonStartPos);
                            }
                            
                            ApplyInsertFormatting(formatRange);
                            System.Diagnostics.Debug.WriteLine($"[操作块格式] 段落{paragraphName}(文档索引{docIndex.Value})添加黄色背景");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用操作块格式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 替换段落内容
        /// </summary>
        private static bool ReplaceParagraph(Word.Document doc, int paragraphIndex, string newContent)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ReplaceParagraph: 请求索引={paragraphIndex}, 文档段落总数={doc.Paragraphs.Count}");

                if (paragraphIndex < 1 || paragraphIndex > doc.Paragraphs.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] ReplaceParagraph: 段落索引无效，索引={paragraphIndex}, 总数={doc.Paragraphs.Count}");
                    return false;
                }

                var paragraph = doc.Paragraphs[paragraphIndex];
                string originalText = paragraph.Range.Text?.TrimEnd('\r', '\n') ?? "";

                System.Diagnostics.Debug.WriteLine($"[DEBUG] ReplaceParagraph: 正在处理段落 {paragraphIndex}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ReplaceParagraph: 原文本长度={originalText.Length}, 内容预览='{originalText.Substring(0, Math.Min(50, originalText.Length))}'");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ReplaceParagraph: 段落范围 - 开始:{paragraph.Range.Start}, 结束:{paragraph.Range.End}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ReplaceParagraph: 新内容长度={(newContent ?? "").Length}");

                // 先给原文本添加删除线格式
                if (!string.IsNullOrEmpty(originalText))
                {
                    ApplyDeleteFormatting(paragraph.Range);
                }

                // 在原文本后面插入新文本（带黄色背景）
                if (!string.IsNullOrEmpty(newContent))
                {
                    Word.Range insertRange = paragraph.Range.Duplicate;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    insertRange.Text = newContent + "\r";
                    ApplyInsertFormatting(insertRange);
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] ReplaceParagraph: 段落 {paragraphIndex} 替换完成");
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 删除段落
        /// </summary>
        private static bool DeleteParagraph(Word.Document doc, int paragraphIndex)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteParagraph: 请求索引={paragraphIndex}, 文档段落总数={doc.Paragraphs.Count}");

                if (paragraphIndex < 1 || paragraphIndex > doc.Paragraphs.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteParagraph: 段落索引无效，索引={paragraphIndex}, 总数={doc.Paragraphs.Count}");
                    return false;
                }

                var paragraph = doc.Paragraphs[paragraphIndex];
                string originalText = paragraph.Range.Text?.TrimEnd('\r', '\n') ?? "";

                System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteParagraph: 正在处理段落 {paragraphIndex}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteParagraph: 原文本长度={originalText.Length}, 内容预览='{originalText.Substring(0, Math.Min(50, originalText.Length))}'");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteParagraph: 段落范围 - 开始:{paragraph.Range.Start}, 结束:{paragraph.Range.End}");

                // 对所有段落（包括空段落）添加删除线标记，不改变文本内容
                ApplyDeleteFormatting(paragraph.Range);

                System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteParagraph: 段落 {paragraphIndex} 删除标记完成");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] DeleteParagraph: 异常 - {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 插入段落
        /// </summary>
        private static bool InsertParagraph(Word.Document doc, int paragraphIndex, string content, string position)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] InsertParagraph: 请求索引={paragraphIndex}, 位置={position}, 文档段落总数={doc.Paragraphs.Count}");

                if (paragraphIndex < 1 || paragraphIndex > doc.Paragraphs.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] InsertParagraph: 段落索引无效，索引={paragraphIndex}, 总数={doc.Paragraphs.Count}");
                    return false;
                }

                var targetParagraph = doc.Paragraphs[paragraphIndex];
                Word.Range insertRange;

                System.Diagnostics.Debug.WriteLine($"[DEBUG] InsertParagraph: 正在处理段落 {paragraphIndex}, 位置={position}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] InsertParagraph: 插入内容长度={(content ?? "").Length}");

                if (position?.ToLower() == "before")
                {
                    // 在段落前插入
                    insertRange = targetParagraph.Range;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }
                else
                {
                    // 在段落后插入（默认）
                    insertRange = targetParagraph.Range;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                insertRange.Text = (content ?? "") + "\r";
                // 给插入的内容加黄色背景
                ApplyInsertFormatting(insertRange);

                System.Diagnostics.Debug.WriteLine($"[DEBUG] InsertParagraph: 段落 {paragraphIndex} 插入完成");
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 应用删除格式（删除线）
        /// </summary>
        private static void ApplyDeleteFormatting(Word.Range range)
        {
            try
            {
                // 创建一个只包含文本内容（不包括段落标记）的范围
                Word.Range contentRange = GetContentOnlyRange(range);
                if (contentRange != null)
                {
                    contentRange.Font.StrikeThrough = 1;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用删除格式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用插入格式（黄色背景）
        /// </summary>
        private static void ApplyInsertFormatting(Word.Range range)
        {
            try
            {
                // 创建一个只包含文本内容（不包括段落标记）的范围
                Word.Range contentRange = GetContentOnlyRange(range);
                if (contentRange != null)
                {
                    contentRange.Shading.BackgroundPatternColor = ConfigManager.DocumentColorHelper.InsertBackground;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用插入格式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取只包含文本内容的范围（不包括段落标记和后续空白）
        /// </summary>
        private static Word.Range GetContentOnlyRange(Word.Range range)
        {
            try
            {
                if (range == null)
                    return null;

                string rangeText = range.Text ?? "";
                
                // 如果范围末尾有段落标记（\r或\n），排除它
                if (rangeText.EndsWith("\r") || rangeText.EndsWith("\n"))
                {
                    // 创建一个新的范围，排除最后一个字符（段落标记）
                    int endPos = range.End - 1;
                    if (endPos >= range.Start)
                    {
                        return range.Document.Range(range.Start, endPos);
                    }
                }
                
                return range;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取内容范围失败: {ex.Message}");
                return range; // 如果出错，返回原范围
            }
        }

        /// <summary>
        /// 段落操作类
        /// </summary>
        private class ParagraphAction
        {
            public string Type { get; set; }
            public int TargetIndex { get; set; } // 段落序号（在ProcessActionsForSnapshot中通过start/end文本查找后设置）
            public string Position { get; set; }
            public string Content { get; set; }
            
            // 用于通过文本定位段落的字段
            public string StartText { get; set; } // replace和delete操作使用
            public string EndText { get; set; } // replace和delete操作使用
            public string PreText { get; set; } // insert操作使用
            public string PostText { get; set; } // insert操作使用
        }

        /// <summary>
        /// 从段落内容中去除接受和丢弃按钮文本
        /// 通过检查书签来识别按钮，一次性去除 " 接受  丢弃 " 这两个按钮
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="paraRange">段落范围</param>
        /// <param name="content">段落内容</param>
        /// <returns>去除按钮文本后的内容</returns>
        private static string RemoveButtonTextFromContent(Word.Document doc, Word.Range paraRange, string content)
        {
            if (string.IsNullOrEmpty(content))
                return content;

            try
            {
                // 检查段落范围中是否包含书签（以 "operation_" 开头的书签）
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    if (bookmark.Name.StartsWith("operation_"))
                    {
                        Word.Range bookmarkRange = bookmark.Range;
                        // 检查书签是否在段落范围内
                        if (bookmarkRange.Start >= paraRange.Start && bookmarkRange.End <= paraRange.End)
                        {
                            // 获取书签在段落中的相对位置
                            int bookmarkStartInPara = bookmarkRange.Start - paraRange.Start;
                            int bookmarkEndInPara = bookmarkRange.End - paraRange.Start;
                            
                            // 确保索引在内容范围内
                            if (bookmarkStartInPara >= 0 && bookmarkEndInPara <= content.Length)
                            {
                                // 从内容中去除书签范围内的文本
                                content = content.Substring(0, bookmarkStartInPara) + content.Substring(bookmarkEndInPara);
                                // 只处理第一个匹配的书签（每个段落通常只有一个按钮组）
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 如果书签检查失败，尝试通过正则表达式匹配 " 接受  丢弃 " 模式
                System.Diagnostics.Debug.WriteLine($"[Hunk删除调试] 书签检查失败，尝试正则匹配: {ex.Message}");
                // 使用正则表达式匹配 " 接受 " 和 " 丢弃 " 连在一起的模式
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"\s*接受\s+丢弃\s*");
                content = regex.Replace(content, "");
            }

            return content;
        }

        /// <summary>
        /// 清除所有操作块（格式和按钮）
        /// 通过书签获取操作范围，精确清除格式
        /// </summary>
        /// <param name="doc">Word文档</param>
        private static void ClearOperationBlocks(Word.Document doc)
        {
            try
            {
                // 第一步：在清除按钮之前，先通过书签清除格式
                System.Diagnostics.Debug.WriteLine("[清除操作块] 第一步：通过书签清除格式");
                var documentProcessor = new DocumentProcessor(doc.Application);
                ClearFormattingByBookmarks(doc, documentProcessor);

                // 第二步：清除所有按钮
                System.Diagnostics.Debug.WriteLine("[清除操作块] 第二步：清除所有按钮");
                int removedButtons = documentProcessor.RemoveAllPendingButtons();
                System.Diagnostics.Debug.WriteLine($"[清除操作块] 已清除 {removedButtons} 个按钮");
                System.Diagnostics.Debug.WriteLine("[清除操作块] 格式清除完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[清除操作块] 清除操作块失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 通过书签获取操作范围，精确清除格式（删除线和背景色）
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="documentProcessor">文档处理器实例</param>
        private static void ClearFormattingByBookmarks(Word.Document doc, DocumentProcessor documentProcessor)
        {
            try
            {
                int strikeThroughCount = 0;
                int backgroundCount = 0;
                
                // 获取所有操作信息
                var operationInfos = DocumentProcessor.OperationInfos;
                
                // 遍历所有书签，查找操作相关的书签
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    string bookmarkName = bookmark.Name;
                    if (!bookmarkName.StartsWith("operation_"))
                        continue;
                    
                    // 从操作信息字典中获取操作详情
                    if (!operationInfos.ContainsKey(bookmarkName))
                        continue;
                    
                    try
                    {
                        dynamic operationInfo = operationInfos[bookmarkName];
                        if (operationInfo == null)
                            continue;
                        
                        string operationType = operationInfo.OperationType;
                        Word.Range originalRange = operationInfo.OriginalRange;
                        Word.Range newRange = operationInfo.NewRange;
                        
                        System.Diagnostics.Debug.WriteLine($"[清除格式] 处理书签 {bookmarkName}, 操作类型: {operationType}");
                        
                        // 根据操作类型清除对应范围的格式
                        switch (operationType?.ToLower())
                        {
                            case "replace":
                                // 替换操作：清除originalRange的删除线，清除newRange的黄色背景
                                if (originalRange != null)
                                {
                                    try
                                    {
                                        Word.Range originalContentRange = GetContentOnlyRange(originalRange);
                                        if (originalContentRange != null)
                                        {
                                            originalContentRange.Font.StrikeThrough = 0;
                                            strikeThroughCount++;
                                            System.Diagnostics.Debug.WriteLine($"[清除格式] 已清除originalRange删除线: {originalRange.Start}-{originalRange.End}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[清除格式] 清除originalRange删除线失败: {ex.Message}");
                                    }
                                }
                                if (newRange != null)
                                {
                                    try
                                    {
                                        // 尝试使用存储的newRange
                                        Word.Range newContentRange = GetContentOnlyRange(newRange);
                                        if (newContentRange != null && newContentRange.Text.Length > 0)
                                        {
                                            // 清除背景色：遍历字符确保清除
                                            for (int i = 1; i <= newContentRange.Characters.Count; i++)
                                            {
                                                var charRange = newContentRange.Characters[i];
                                                if (charRange.Shading.BackgroundPatternColor == ConfigManager.DocumentColorHelper.InsertBackground)
                                                {
                                                    charRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                                                }
                                            }
                                            backgroundCount++;
                                            System.Diagnostics.Debug.WriteLine($"[清除格式] 已清除newRange黄色背景: {newRange.Start}-{newRange.End}");
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[清除格式] newRange无效或为空: Start={newRange?.Start}, End={newRange?.End}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[清除格式] 清除newRange黄色背景失败: {ex.Message}");
                                        // 如果newRange失效，尝试通过书签位置重新获取
                                        try
                                        {
                                            // 书签名称格式: operation_replace_{start}_{end}
                                            // 书签本身的范围就是按钮的位置，我们需要找到对应的文本范围
                                            // 由于newRange可能已经失效，我们需要通过其他方式找到黄色背景的文本
                                        }
                                        catch { }
                                    }
                                }
                                break;
                            
                            case "delete":
                                // 删除操作：清除originalRange的删除线
                                if (originalRange != null)
                                {
                                    try
                                    {
                                        Word.Range originalContentRange = GetContentOnlyRange(originalRange);
                                        if (originalContentRange != null)
                                        {
                                            originalContentRange.Font.StrikeThrough = 0;
                                            strikeThroughCount++;
                                            System.Diagnostics.Debug.WriteLine($"[清除格式] 已清除originalRange删除线: {originalRange.Start}-{originalRange.End}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[清除格式] 清除originalRange删除线失败: {ex.Message}");
                                    }
                                }
                                break;
                            
                            case "insert":
                                // 插入操作：清除newRange的黄色背景
                                if (newRange != null)
                                {
                                    try
                                    {
                                        Word.Range newContentRange = GetContentOnlyRange(newRange);
                                        if (newContentRange != null && newContentRange.Text.Length > 0)
                                        {
                                            // 清除背景色：遍历字符确保清除
                                            for (int i = 1; i <= newContentRange.Characters.Count; i++)
                                            {
                                                var charRange = newContentRange.Characters[i];
                                                if (charRange.Shading.BackgroundPatternColor == ConfigManager.DocumentColorHelper.InsertBackground)
                                                {
                                                    charRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                                                }
                                            }
                                            backgroundCount++;
                                            System.Diagnostics.Debug.WriteLine($"[清除格式] 已清除newRange黄色背景: {newRange.Start}-{newRange.End}");
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[清除格式] newRange无效或为空: Start={newRange?.Start}, End={newRange?.End}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[清除格式] 清除newRange黄色背景失败: {ex.Message}");
                                    }
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[清除格式] 处理书签 {bookmarkName} 时出错: {ex.Message}");
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[清除格式] 已清除 {strikeThroughCount} 处删除线");
                System.Diagnostics.Debug.WriteLine($"[清除格式] 已清除 {backgroundCount} 处黄色背景");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[清除格式] 通过书签清除格式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查文档中是否还有待处理的按钮
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <returns>如果还有按钮存在，返回true；否则返回false</returns>
        private static bool HasRemainingButtons(Word.Document doc)
        {
            try
            {
                // 检查是否有以 "operation_" 开头的书签（这些是按钮标记）
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    if (bookmark.Name.StartsWith("operation_"))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[清除操作块] 检查按钮时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 操作结果类
        /// </summary>
        private class ActionResult
        {
            public int action_index { get; set; }
            public string type { get; set; }
            public int target_index { get; set; }
            public bool success { get; set; }
            public string message { get; set; }
        }
    }
}
