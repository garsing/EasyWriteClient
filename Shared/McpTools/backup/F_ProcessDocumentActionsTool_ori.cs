using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 处理句子操作工具
    /// 支持通过文本定位对Word文档进行句子级别的替换、删除、插入等操作
    /// </summary>
    public static class F_ProcessDocumentActionsTool
    {
        /// <summary>
        /// 注册处理句子操作工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            List<McpTool> availableTools,
            object wordApplication)
        {
            var tool = new McpTool
            {
                name = "F_process_document_actions",
                alias = "处理句子操作",
                description = "通过句子编号序列对Word文档进行批量句子操作，支持替换、删除、插入句子内容。。通过句子编号序列（如AAAAA,AAAAB,AAAAC）在快照中定位要操作的句子。调用格式：{\"action\": [{\"type\": \"replace\", \"detail\": {\"original\": {\"sentence_names\": \"AAAAA,AAAAB,AAAAC\"(必需，句子编号序列，逗号分隔)}, \"new\": \"新内容\"(必需，必须是完整句子，会自动切分为句子)}}, {\"type\": \"delete\", \"detail\": {\"deletion\": {\"sentence_names\": \"AAAAA,AAAAB\"(必需，句子编号序列)}}}, {\"type\": \"insert\", \"detail\": {\"pre\": {\"sentence_names\": \"AAAAA\"(必需，上文句子编号序列)}, \"insert\": \"插入内容\"(必需，必须是完整句子，会自动切分为句子), \"post\": {\"sentence_names\": \"AAAAB\"(必需，下文句子编号序列)}}}}]}。替换显示原文本(删除线)+新文本(黄色背景)，删除保留原文本(删除线)，插入显示新文本(黄色背景)。注意：使用前请先调用F_get_document_content工具获取当前文档的句子编号序列，然后使用这些编号来指定要操作的句子。句子编号序列必须在快照中存在且唯一，否则操作会失败。⚠️⚠️⚠️重要：1. \"sentence_names\"必须是在文本中连续的序列，F_get_document_content读到的文本是{\"AAAAA\":\"xxxxxxx\",\"AAABD\":\"xxxxxxx\",\"AAAEE\":\"xxxxxxx\",\"AAAQA\":\"xxxxxxx\",\"ABAQA\":\"xxxxxxx\"}，如果你想改AAAAA,AAABD,AAAQA,ABAQA就不能写{\"action\": [{\"type\": \"replace\",\"detail\": {\"original\": {\"sentence_names\": \"AAAAA,AAABD,AAAQA,ABAQA\"},\"new\": \"新内容\"}}]}，因为AAAAA,AAABD,AAAQA,ABAQA不连续，中间隔了AAAEE，你必须分开写{\"action\": [{\"type\": \"replace\",\"detail\": {\"original\": {\"sentence_names\": \"AAAAA,AAABD\"},\"new\": \"新内容1\"}},{\"type\": \"replace\",\"detail\": {\"original\": {\"sentence_names\": \"AAAQA,ABAQA\"},\"new\": \"新内容2\"}}]} 2. 你在替换新文本的时候如果你觉得要分段一定要在你觉得需要的地方加上换行符\\r",
                usage = @"批量修改文档句子内容。

⚠️⚠️⚠️重要限制说明：
本工具是以句子为级别的修改工具，所有操作都是针对完整句子进行的。

⚠️必须遵守的规则：
1. 修改的新内容必须是完整句子（不能只修改句子的一部分）
   - 例如：new内容应该是""人工智能是计算机科学的重要分支。""（完整句子）
   - ❌ 错误：只提供""人工智能是计算机科学""（不完整，缺少句号）

2. ⚠️无法只修改句子中的一部分，如果想修改句子中的一部分，必须将整个句子进行修改
   - 例如：想修改""人工智能是计算机科学的一个分支。""中的""一个""为""重要""
   - ✅ 正确做法：将整个句子替换为""人工智能是计算机科学的重要分支。""
   - ❌ 错误做法：只替换""一个""为""重要""

⚠️工作原理：
- 使用句子编号序列（如AAAAA,AAAAB,AAAAC）在快照中定位要操作的句子
- 句子编号序列必须在快照中存在且唯一，否则操作会失败
- 对new内容切分为句子，为每个句子创建名字
- 执行替换、删除或插入操作

⚠️使用步骤：
1. 首先调用F_get_document_content工具获取当前文档的句子编号序列
2. 从返回结果中找到要操作的句子编号（格式：AAAAA:句子内容）
3. 使用这些句子编号来指定要操作的句子

支持三种操作：

1) replace-替换句子：
   - 必需字段：original.sentence_names（句子编号序列，如""AAAAA,AAAAB,AAAAC""，逗号分隔）、new（新内容，必须是完整句子，会自动切分为句子）
   - ⚠️注意：sentence_names指定的句子编号序列必须在快照中存在且唯一，new内容也必须是完整句子
   - 通过句子编号序列找到快照位置，然后用new内容切分的句子替换
   - 显示效果：原文本(删除线)+新文本(黄色背景)

2) delete-删除句子：
   - 必需字段：deletion.sentence_names（句子编号序列，如""AAAAA,AAAAB""）
   - ⚠️注意：sentence_names指定的句子编号序列必须在快照中存在且唯一
   - 通过句子编号序列找到快照位置，然后删除这些句子
   - 显示效果：保留原文本并添加删除线标记

3) insert-插入句子：
   - 必需字段：pre.sentence_names（上文句子编号序列，如""AAAAA""）、insert（要插入的内容，必须是完整句子，会自动切分为句子）、post.sentence_names（下文句子编号序列，如""AAAAB""）
   - ⚠️注意：pre和post的句子编号序列必须在快照中存在且唯一，insert内容必须是完整句子
   - 通过pre和post的句子编号序列定位插入位置（在pre序列之后、post序列之前），然后插入insert内容切分的句子
   - 显示效果：新文本(黄色背景)

句子切分规则：
本工具使用以下标点符号作为句子分隔符：
- 中文标点：句号(。)、分号(；)、问号(？)、感叹号(！)、冒号(：)
- 英文标点：句号(.)、分号(;)、问号(?)、感叹号(!)、冒号(:)

当文本遇到以上任一标点符号时，会在该标点符号后切分为新的句子。",
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

            toolRegistry["F_process_document_actions"] = async (args) =>
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

                    // 打印操作前文档内容用于诊断，并直接切分为句子（不先划分段落）
                    var sentencesBefore = new List<string>();
                    // 获取整个文档内容（保留换行符）
                    string fullContent = wordApp.ActiveDocument.Content.Text ?? "";
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作前文档内容长度: {fullContent.Length}");
                    
                    // 直接对整个文档内容切分为句子
                    List<string> sentences = SplitIntoSentences(fullContent);
                        foreach (var sentence in sentences)
                        {
                            // 跳过完全空白的句子，但保留只包含换行符的句子
                            if (!string.IsNullOrEmpty(sentence) && !(sentence.Trim().Length == 0 && !sentence.Contains("\r") && !sentence.Contains("\n")))
                            {
                                sentencesBefore.Add(sentence);
                            }
                        }
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作前句子数量: {sentencesBefore.Count}");

                    // 第二步：初始化句子跟踪（如果还没有快照的话）
                    System.Diagnostics.Debug.WriteLine("[第二步] 检查句子跟踪状态");
                    var currentSnapshot = SentenceState.CurrentSnapshot;
                    if (currentSnapshot == null || currentSnapshot.Count == 0)
                    {
                        // 没有快照，需要初始化
                        System.Diagnostics.Debug.WriteLine("[第二步] 初始化句子跟踪（首次调用）");
                        SentenceState.InitializeSentenceTracking(sentencesBefore);
                    }
                    else
                    {
                        // 已有快照，基于最新快照继续（不重新初始化）
                        System.Diagnostics.Debug.WriteLine($"[第二步] 已有快照，基于最新快照继续（快照数量: {SentenceState.Snapshots.Count}）");
                    }
                    System.Diagnostics.Debug.WriteLine("[第二步完成]");

                    if (!args.ContainsKey("action"))
                    {
                        return new ToolResult { Success = false, Error = "缺少必需参数：action" };
                    }

                    var parseResult = ParseActions(args["action"], wordApp.ActiveDocument);
                    bool parseSuccess = parseResult.Item1;
                    string parseError = parseResult.Item2;
                    List<ParagraphAction> actions = parseResult.Item3;
                    
                    if (!parseSuccess)
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }
                    
                    // 第三步：生成新的快照（不操作文档）
                    System.Diagnostics.Debug.WriteLine("[第三步] 生成新的快照（不操作文档）");
                    string errorMessage = await ProcessActionsForSnapshot(actions, wordApp.ActiveDocument);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        // 出现异常，恢复按钮和格式：运行第六步根据display_hunk修改格式
                        System.Diagnostics.Debug.WriteLine("[异常恢复] 出现异常，恢复按钮和格式");
                        var restoreDisplayHunk = SentenceState.DisplayHunk;
                        if (restoreDisplayHunk != null && restoreDisplayHunk.Count > 0)
                        {
                            // 将display_hunk转换为hunk格式（List<(string header, List<string> lines)>）
                            var restoreHunk = new List<(string header, List<string> lines)>();
                            foreach (var (header, lines) in restoreDisplayHunk)
                            {
                                restoreHunk.Add((header, lines.ToList()));
                            }
                            
                            // 获取targetFullDisplayContent
                            var restoreTargetFullDisplayContent = SentenceState.GetTargetFullDisplayContent(restoreHunk);
                            
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
                    var hunk = SentenceState.CalculateAndPrintHunk();
                    System.Diagnostics.Debug.WriteLine($"[第四步完成] 共生成 {hunk.Count} 个全量hunk（包含所有句子）");
                    
                    // 打印句子名称到内容的对应关系（打印整个映射表，包括所有历史映射）
                    System.Diagnostics.Debug.WriteLine("[句子映射表] ==========");
                    var allMappings = SentenceState.GetAllSentenceMappings();
                    System.Diagnostics.Debug.WriteLine($"[句子映射表] 映射表总数: {allMappings.Count}");
                    for (int i = 0; i < allMappings.Count; i++)
                    {
                        var mapping = allMappings[i];
                        // 将换行符转换为可见字符以便调试输出
                        string displayContent = mapping.content.Replace("\r\n", "\\r\\n").Replace("\r", "\\r").Replace("\n", "\\n");
                        System.Diagnostics.Debug.WriteLine($"[句子映射表] 映射{i + 1} -> {mapping.name}: {displayContent}");
                    }
                    System.Diagnostics.Debug.WriteLine("[句子映射表] ==========");

                    // 第五步：根据全量hunk执行文档插入操作（只处理+号的句子）
                    System.Diagnostics.Debug.WriteLine("[第五步] 根据全量hunk执行文档插入操作");
                    // 获取起点（originalFullDisplayContent）：从当前的display_hunk直接获取（全量形式）
                    var originalFullDisplayContent = SentenceState.GetOriginalFullDisplayContent();
                    System.Diagnostics.Debug.WriteLine($"[第五步] originalFullDisplayContent（起点）: {string.Join(",", originalFullDisplayContent)}");
                    
                    // 获取终点（targetFullDisplayContent）：从新的全量hunk直接获取
                    var targetFullDisplayContent = SentenceState.GetTargetFullDisplayContent(hunk);
                    System.Diagnostics.Debug.WriteLine($"[第五步] targetFullDisplayContent（终点）: {string.Join(",", targetFullDisplayContent)}");
                    
                    // 对比起点和终点，找出需要插入的句子
                    await ApplyHunkInsertions(wordApp.ActiveDocument, hunk, originalFullDisplayContent, targetFullDisplayContent);
                    System.Diagnostics.Debug.WriteLine("[第五步完成]");

                    // 第五步-2：更新display_hunk
                    System.Diagnostics.Debug.WriteLine("[第五步-2] 更新display_hunk");
                    
                    // 打印全量的名字句子内容映射表
                    System.Diagnostics.Debug.WriteLine("[第五步-2] ========== 全量名字句子内容映射表 ==========");
                    var allMappingsForStep5 = SentenceState.GetAllSentenceMappings();
                    System.Diagnostics.Debug.WriteLine($"[第五步-2] 映射表总数: {allMappingsForStep5.Count}");
                    foreach (var mapping in allMappingsForStep5)
                    {
                        // 将换行符转换为可见字符以便调试输出
                        string displayContent = mapping.content.Replace("\r\n", "\\r\\n").Replace("\r", "\\r").Replace("\n", "\\n");
                        System.Diagnostics.Debug.WriteLine($"[第五步-2] {mapping.name} -> {displayContent}");
                    }
                    System.Diagnostics.Debug.WriteLine("[第五步-2] ==========================================");
                    
                    // 直接存储全量hunk列表（包含所有句子）
                    SentenceState.UpdateDisplayHunk(hunk);

                    // 打印更新后的display_hunk用于调试
                    var displayHunk = SentenceState.DisplayHunk;
                    System.Diagnostics.Debug.WriteLine($"[第五步-2] display_hunk更新完成，当前包含 {displayHunk.Count} 个全量hunk:");
                    foreach (var (header, lines) in displayHunk)
                    {
                        System.Diagnostics.Debug.WriteLine($"  全量Hunk (header为空): 包含 {lines.Count} 行句子");
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
                        // 全量hunk：直接使用hunkLines（包含所有句子）
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
                    return new ToolResult { Success = false, Error = $"处理句子操作失败: {ex.Message}" };
                }
            };

            availableTools.Add(tool);
        }

        /// <summary>
        /// 解析操作列表（新格式：存储句子编号序列，稍后在ProcessActionsForSnapshot中通过句子编号序列在快照中查找）
        /// </summary>
        private static (bool success, string error, List<ParagraphAction> actions) ParseActions(object actionsObj, Word.Document doc)
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
                                // 解析replace操作：original.sentence_names（句子编号序列，如 "AAAAA,AAAAB,AAAAC"）, new
                                object original = detail.ContainsKey("original") ? detail["original"] : null;
                                if (original is Dictionary<string, object> originalDict)
                                {
                                    action.SentenceNames = originalDict.ContainsKey("sentence_names") ? originalDict["sentence_names"]?.ToString() : null;
                                }
                                
                                // 解析新内容
                                if (detail.ContainsKey("new"))
                                {
                                    action.Content = detail["new"]?.ToString();
                                }
                                break;

                            case "delete":
                                // 解析delete操作：deletion.sentence_names（句子编号序列）
                                object deletion = detail.ContainsKey("deletion") ? detail["deletion"] : null;
                                if (deletion is Dictionary<string, object> deletionDict)
                                {
                                    action.SentenceNames = deletionDict.ContainsKey("sentence_names") ? deletionDict["sentence_names"]?.ToString() : null;
                                }
                                break;

                            case "insert":
                                // 解析insert操作：pre.sentence_names, post.sentence_names, insert
                                object pre = detail.ContainsKey("pre") ? detail["pre"] : null;
                                if (pre is Dictionary<string, object> preDict)
                                {
                                    action.PreSentenceNames = preDict.ContainsKey("sentence_names") ? preDict["sentence_names"]?.ToString() : null;
                                }
                                
                                object post = detail.ContainsKey("post") ? detail["post"] : null;
                                if (post is Dictionary<string, object> postDict)
                                {
                                    action.PostSentenceNames = postDict.ContainsKey("sentence_names") ? postDict["sentence_names"]?.ToString() : null;
                                }
                                
                                action.Position = "after"; // 默认在定位句子之后插入
                                
                                // 解析插入内容
                                if (detail.ContainsKey("insert"))
                                {
                                    action.Content = detail["insert"]?.ToString();
                                }
                                break;
                        }

                        // 验证必要字段并检查句子序列连续性
                        bool isValid = false;
                        string continuityError = null;
                        
                        if (action.Type?.ToLower() == "replace")
                        {
                            // replace操作：需要sentence_names和new内容
                            isValid = !string.IsNullOrEmpty(action.SentenceNames) && !string.IsNullOrEmpty(action.Content);
                            if (isValid && !string.IsNullOrEmpty(action.SentenceNames))
                            {
                                continuityError = ValidateSentenceSequenceContinuity(action.SentenceNames, action.Type);
                            }
                        }
                        else if (action.Type?.ToLower() == "delete")
                        {
                            // delete操作：需要sentence_names
                            isValid = !string.IsNullOrEmpty(action.SentenceNames);
                            if (isValid && !string.IsNullOrEmpty(action.SentenceNames))
                            {
                                continuityError = ValidateSentenceSequenceContinuity(action.SentenceNames, action.Type);
                            }
                        }
                        else if (action.Type?.ToLower() == "insert")
                        {
                            // insert操作：需要pre.sentence_names或post.sentence_names，以及insert内容
                            isValid = (!string.IsNullOrEmpty(action.PreSentenceNames) || !string.IsNullOrEmpty(action.PostSentenceNames)) && !string.IsNullOrEmpty(action.Content);
                            if (isValid)
                            {
                                // 检查pre和post的连续性（如果存在）
                                if (!string.IsNullOrEmpty(action.PreSentenceNames))
                                {
                                    continuityError = ValidateSentenceSequenceContinuity(action.PreSentenceNames, action.Type + " (pre)");
                                    if (continuityError != null) isValid = false;
                                }
                                if (isValid && !string.IsNullOrEmpty(action.PostSentenceNames))
                                {
                                    continuityError = ValidateSentenceSequenceContinuity(action.PostSentenceNames, action.Type + " (post)");
                                    if (continuityError != null) isValid = false;
                                }
                            }
                        }
                        
                        // 如果连续性检查失败，返回错误
                        if (!string.IsNullOrEmpty(continuityError))
                        {
                            return (false, continuityError, null);
                        }

                        if (isValid)
                        {
                            actions.Add(action);
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 解析操作失败：缺少必要字段，操作类型={action.Type}, SentenceNames={action.SentenceNames}, PreSentenceNames={action.PreSentenceNames}, PostSentenceNames={action.PostSentenceNames}, Content={action.Content}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 解析操作失败: {ex.Message}");
                return (false, $"解析操作失败: {ex.Message}", null);
            }

            return (true, null, actions);
        }

        /// <summary>
        /// 通过文本范围找到段落索引（用于定位段落，然后从段落中提取句子）
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
        /// 句子信息（包含内容和位置）
        /// </summary>
        private class SentenceInfo
        {
            public string Text { get; set; }
            public int StartIndex { get; set; }  // 在文本中的起始位置（从0开始）
            public int EndIndex { get; set; }    // 在文本中的结束位置（从0开始，包含）
        }

        /// <summary>
        /// 切分文本为句子列表（使用句号、分号、问号、感叹号、冒号等作为分隔符）
        /// 分隔符包括：中文标点（。；？！：）和英文标点（.;?!:）
        /// </summary>
        /// <param name="text">要切分的文本</param>
        /// <returns>句子列表</returns>
        private static List<string> SplitIntoSentences(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<string>();

            var sentences = new List<string>();
            var sentenceEndings = new[] { '。', '；', '？', '！', '：', '.', ';', '?', '!', ':' };
            
            int start = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (sentenceEndings.Contains(text[i]))
                {
                    // 找到句子结束符，检查后面是否有换行符
                    int sentenceEnd = i + 1; // 句子结束符的位置+1
                    
                    // 跳过句子结束符后的空格和制表符
                    while (sentenceEnd < text.Length && (text[sentenceEnd] == ' ' || text[sentenceEnd] == '\t'))
                    {
                        sentenceEnd++;
                    }
                    
                    // 如果后面有换行符，包含它
                    if (sentenceEnd < text.Length && (text[sentenceEnd] == '\r' || text[sentenceEnd] == '\n'))
                    {
                        sentenceEnd++;
                        // 如果是\r\n组合，包含两个字符
                        if (sentenceEnd < text.Length && text[sentenceEnd - 1] == '\r' && text[sentenceEnd] == '\n')
                        {
                            sentenceEnd++;
                        }
                    }
                    
                    // 提取句子（包含结束符和可能的换行符）
                    // 只去掉前面的空格和制表符，保留换行符
                    int actualStart = start;
                    while (actualStart < text.Length && (text[actualStart] == ' ' || text[actualStart] == '\t'))
                    {
                        actualStart++;
                    }
                    
                    string sentence = text.Substring(actualStart, sentenceEnd - actualStart);
                    if (!string.IsNullOrEmpty(sentence))
                    {
                        sentences.Add(sentence);
                    }
                    start = sentenceEnd;
                    i = sentenceEnd - 1; // 调整i的位置，因为已经处理了换行符
                }
            }
            
            // 处理最后一段（如果没有以句子结束符结尾）
            if (start < text.Length)
            {
                // 只去掉前面的空格和制表符，保留换行符
                int actualStart = start;
                while (actualStart < text.Length && (text[actualStart] == ' ' || text[actualStart] == '\t'))
                {
                    actualStart++;
                }
                
                string lastSentence = text.Substring(actualStart);
                if (!string.IsNullOrEmpty(lastSentence))
                {
                    sentences.Add(lastSentence);
                }
            }
            
            return sentences;
        }

        /// <summary>
        /// 切分文本为句子信息列表（包含位置信息）
        /// </summary>
        /// <param name="text">要切分的文本</param>
        /// <returns>句子信息列表（包含内容和位置）</returns>
        private static List<SentenceInfo> SplitIntoSentencesWithPosition(string text)
        {
            var result = new List<SentenceInfo>();
            if (string.IsNullOrWhiteSpace(text))
                return result;

            var sentenceEndings = new[] { '。', '；', '？', '！', '：', '.', ';', '?', '!', ':' };
            
            int start = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (sentenceEndings.Contains(text[i]))
                {
                    // 找到句子结束符，检查后面是否有换行符
                    int sentenceEnd = i + 1; // 句子结束符的位置+1
                    
                    // 跳过句子结束符后的空格和制表符
                    while (sentenceEnd < text.Length && (text[sentenceEnd] == ' ' || text[sentenceEnd] == '\t'))
                    {
                        sentenceEnd++;
                    }
                    
                    // 如果后面有换行符，包含它
                    if (sentenceEnd < text.Length && (text[sentenceEnd] == '\r' || text[sentenceEnd] == '\n'))
                    {
                        sentenceEnd++;
                        // 如果是\r\n组合，包含两个字符
                        if (sentenceEnd < text.Length && text[sentenceEnd - 1] == '\r' && text[sentenceEnd] == '\n')
                        {
                            sentenceEnd++;
                        }
                    }
                    
                    // 提取句子（包含结束符和可能的换行符）
                    // 只去掉前面的空格和制表符，保留换行符
                    int actualStart = start;
                    while (actualStart < text.Length && (text[actualStart] == ' ' || text[actualStart] == '\t'))
                    {
                        actualStart++;
                    }
                    
                    string sentence = text.Substring(actualStart, sentenceEnd - actualStart);
                    if (!string.IsNullOrEmpty(sentence))
                    {
                        result.Add(new SentenceInfo
                        {
                            Text = sentence,
                            StartIndex = actualStart,
                            EndIndex = sentenceEnd - 1
                        });
                    }
                    start = sentenceEnd;
                    i = sentenceEnd - 1; // 调整i的位置，因为已经处理了换行符
                }
            }
            
            // 处理最后一段（如果没有以句子结束符结尾）
            if (start < text.Length)
            {
                // 只去掉前面的空格和制表符，保留换行符
                int actualStart = start;
                while (actualStart < text.Length && (text[actualStart] == ' ' || text[actualStart] == '\t'))
                {
                    actualStart++;
                    }
                    
                string lastSentence = text.Substring(actualStart);
                if (!string.IsNullOrEmpty(lastSentence))
                {
                    result.Add(new SentenceInfo
                    {
                        Text = lastSentence,
                        StartIndex = actualStart,
                        EndIndex = text.Length - 1
                    });
                }
            }
            
            return result;
        }

        /// <summary>
        /// 处理句子操作
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
                                    // 句子级别操作已在ProcessActionsForSnapshot中处理
                                    // DocumentState.ProcessParagraphReplace(action.TargetIndex, action.Content);
                                }
                                result.message = result.success ? "替换成功" : "替换失败：段落不存在";
                                System.Diagnostics.Debug.WriteLine($"[操作结果] {result.message}");
                                break;

                            case "delete":
                                System.Diagnostics.Debug.WriteLine($"[操作步骤] 执行删除操作：段落{action.TargetIndex}");
                                result.success = DeleteParagraph(doc, action.TargetIndex);
                                if (result.success)
                                {
                                    // 句子级别操作已在ProcessActionsForSnapshot中处理
                                    // DocumentState.ProcessParagraphDelete(action.TargetIndex);
                                }
                                result.message = result.success ? "删除成功" : "删除失败：段落不存在";
                                System.Diagnostics.Debug.WriteLine($"[操作结果] {result.message}");
                                break;

                            case "insert":
                                System.Diagnostics.Debug.WriteLine($"[操作步骤] 执行插入操作：在段落{action.TargetIndex} {action.Position} 插入内容");
                                result.success = InsertParagraph(doc, action.TargetIndex, action.Content, action.Position);
                                if (result.success)
                                {
                                    // 句子级别操作已在ProcessActionsForSnapshot中处理
                                    // DocumentState.ProcessParagraphInsert(action.TargetIndex, action.Position, action.Content);
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
        /// 处理句子操作（只更新快照，不操作文档）
        /// 通过句子编号序列在快照中查找位置进行操作
        /// </summary>
        /// <returns>错误信息，如果成功则返回null或空字符串</returns>
        private static async Task<string> ProcessActionsForSnapshot(List<ParagraphAction> actions, Word.Document doc)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[句子快照] 当前快照包含 {SentenceState.CurrentSnapshot.Count} 个句子");

                // 按原始索引顺序处理操作（不排序），维护索引映射表
                foreach (var action in actions)
                {
                    try
                    {
                        // 每次操作前获取最新的快照
                        var currentSnapshot = SentenceState.CurrentSnapshot.ToList();
                        
                        List<string> originalSentenceNames = new List<string>();
                        int startIndex = -1;
                        
                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                            case "delete":
                                // 通过句子编号序列在快照中查找
                                if (string.IsNullOrEmpty(action.SentenceNames))
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] {action.Type}操作：缺少sentence_names");
                                    return $"{action.Type}操作缺少sentence_names参数，请输入句子编号序列（如 AAAAA,AAAAB,AAAAC）";
                                }
                                
                                // 解析句子编号序列（逗号分隔）
                                originalSentenceNames = action.SentenceNames.Split(',')
                                    .Select(s => s.Trim())
                                    .Where(s => !string.IsNullOrEmpty(s))
                                    .ToList();
                                
                                if (originalSentenceNames.Count == 0)
                                {
                                    return $"句子编号序列格式错误：{action.SentenceNames}。请输入正确的句子编号序列（如 AAAAA,AAAAB,AAAAC）";
                                }
                                
                                System.Diagnostics.Debug.WriteLine($"[快照操作] {action.Type}操作：查找句子编号序列 {string.Join(",", originalSentenceNames)}");
                                
                                // 在快照中查找序列是否存在且唯一
                                List<int> matchPositions = new List<int>();
                                for (int i = 0; i <= currentSnapshot.Count - originalSentenceNames.Count; i++)
                                {
                                    bool match = true;
                                    for (int j = 0; j < originalSentenceNames.Count; j++)
                                    {
                                        if (currentSnapshot[i + j] != originalSentenceNames[j])
                                        {
                                            match = false;
                                            break;
                                        }
                                    }
                                    if (match)
                                    {
                                        matchPositions.Add(i + 1); // 转换为从1开始的索引
                                    }
                                }
                                
                                if (matchPositions.Count == 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 错误：句子序列在快照中不存在");
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 查找的序列: {string.Join(",", originalSentenceNames)}");
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 当前快照长度: {currentSnapshot.Count}");
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 当前快照前10个元素: {string.Join(",", currentSnapshot.Take(10))}");
                                    return $"句子编号序列在快照中不存在：{action.SentenceNames}。请使用F_get_document_content工具查看当前文档的句子编号";
                                }
                                
                                if (matchPositions.Count > 1)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 错误：句子序列在快照中不唯一，找到{matchPositions.Count}个匹配位置");
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 查找的序列: {string.Join(",", originalSentenceNames)}");
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 匹配位置: {string.Join(",", matchPositions.Select(p => $"位置{p}"))}");
                                    for (int idx = 0; idx < matchPositions.Count; idx++)
                                    {
                                        int pos = matchPositions[idx];
                                        int endPos = pos + originalSentenceNames.Count - 1;
                                        System.Diagnostics.Debug.WriteLine($"[快照操作]   匹配{idx + 1}: 位置{pos}-{endPos}, 内容: {string.Join(",", currentSnapshot.Skip(pos - 1).Take(originalSentenceNames.Count))}");
                                    }
                                    return $"句子编号序列在快照中不唯一，找到{matchPositions.Count}个匹配位置。请使用更具体的句子编号序列";
                                }
                                
                                startIndex = matchPositions[0];
                                System.Diagnostics.Debug.WriteLine($"[快照操作] {action.Type}操作：找到唯一匹配位置 {startIndex}");
                                break;

                            case "insert":
                                // 通过pre和post句子编号序列定位插入位置
                                if (string.IsNullOrEmpty(action.PreSentenceNames) || string.IsNullOrEmpty(action.PostSentenceNames))
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 插入操作：缺少pre_sentence_names或post_sentence_names");
                                    return $"插入操作缺少pre.sentence_names或post.sentence_names参数，请输入句子编号序列（如 AAAAA,AAAAB）";
                                }
                                
                                // 解析pre和post句子编号序列
                                var preSentenceNames = action.PreSentenceNames.Split(',')
                                    .Select(s => s.Trim())
                                    .Where(s => !string.IsNullOrEmpty(s))
                                    .ToList();
                                var postSentenceNames = action.PostSentenceNames.Split(',')
                                    .Select(s => s.Trim())
                                    .Where(s => !string.IsNullOrEmpty(s))
                                    .ToList();
                                
                                if (preSentenceNames.Count == 0 || postSentenceNames.Count == 0)
                                {
                                    return $"句子编号序列格式错误：pre={action.PreSentenceNames}, post={action.PostSentenceNames}。请输入正确的句子编号序列";
                                }
                                
                                // 在快照中查找pre序列的结束位置和post序列的开始位置
                                int preEndIndex = -1;
                                int postStartIndex = -1;
                                
                                // 查找pre序列的结束位置（最后一个句子的位置）
                                List<int> preMatchPositions = new List<int>();
                                for (int i = 0; i <= currentSnapshot.Count - preSentenceNames.Count; i++)
                                {
                                    bool match = true;
                                    for (int j = 0; j < preSentenceNames.Count; j++)
                                    {
                                        if (currentSnapshot[i + j] != preSentenceNames[j])
                                        {
                                            match = false;
                                            break;
                                        }
                                    }
                                    if (match)
                                    {
                                        preMatchPositions.Add(i + preSentenceNames.Count); // pre序列结束后的位置
                                    }
                                }
                                
                                // 查找post序列的开始位置
                                List<int> postMatchPositions = new List<int>();
                                for (int i = 0; i <= currentSnapshot.Count - postSentenceNames.Count; i++)
                                {
                                    bool match = true;
                                    for (int j = 0; j < postSentenceNames.Count; j++)
                                    {
                                        if (currentSnapshot[i + j] != postSentenceNames[j])
                                        {
                                            match = false;
                                            break;
                                        }
                                    }
                                    if (match)
                                    {
                                        postMatchPositions.Add(i + 1); // post序列开始的位置（从1开始）
                                    }
                                }
                                
                                if (preMatchPositions.Count == 0)
                                {
                                    return $"pre句子编号序列在快照中不存在：{action.PreSentenceNames}";
                                }
                                if (preMatchPositions.Count > 1)
                                {
                                    return $"pre句子编号序列在快照中不唯一，找到{preMatchPositions.Count}个匹配位置：{action.PreSentenceNames}";
                                }
                                if (postMatchPositions.Count == 0)
                                {
                                    return $"post句子编号序列在快照中不存在：{action.PostSentenceNames}";
                                }
                                if (postMatchPositions.Count > 1)
                                {
                                    return $"post句子编号序列在快照中不唯一，找到{postMatchPositions.Count}个匹配位置：{action.PostSentenceNames}";
                                }
                                
                                preEndIndex = preMatchPositions[0];
                                postStartIndex = postMatchPositions[0];
                                
                                // 验证pre和post的位置关系
                                if (preEndIndex >= postStartIndex)
                                {
                                    return $"pre和post句子编号序列位置关系错误：pre结束位置({preEndIndex})应该小于post开始位置({postStartIndex})";
                                }
                                
                                // 插入位置在pre序列之后
                                startIndex = preEndIndex;
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 插入操作：在pre序列结束位置{preEndIndex}和post序列开始位置{postStartIndex}之间插入");
                                    break;
                        }

                        if (startIndex <= 0 || startIndex > currentSnapshot.Count)
                        {
                            System.Diagnostics.Debug.WriteLine($"[快照操作] 警告：句子序号 {startIndex} 无效，跳过操作");
                            continue;
                        }

                        // 计算endIndex（对于replace和delete操作）
                        int endIndex = startIndex;
                        if (action.Type?.ToLower() == "replace" || action.Type?.ToLower() == "delete")
                        {
                            if (originalSentenceNames.Count == 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 错误：句子编号序列为空");
                                return "句子编号序列为空";
                            }
                            endIndex = startIndex + originalSentenceNames.Count - 1;
                            System.Diagnostics.Debug.WriteLine($"[索引映射] 操作 {action.Type}：找到句子序列位置 {startIndex}-{endIndex}，名字序列: {string.Join(",", originalSentenceNames)}");
                        }

                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                                // 将new内容切分为句子，为每个句子创建名字
                                List<string> newSentences = SplitIntoSentences(action.Content ?? "");
                                if (newSentences.Count == 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 警告：new内容无法切分为句子，跳过操作");
                                    continue;
                                }
                                
                                // 为新句子创建名字并打印调试信息
                                List<string> newSentenceNames = new List<string>();
                                foreach (var sentence in newSentences)
                                {
                                    string name = SentenceState.FindOrCreateSentenceName(sentence);
                                    newSentenceNames.Add(name);
                                }
                                
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 替换操作：句子{startIndex}-{endIndex} -> {newSentences.Count}个新句子");
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 新句子的名字: {string.Join(",", newSentenceNames)}");
                                SentenceState.ProcessSentenceReplace(startIndex, endIndex, newSentences);
                                // 替换操作后，快照已更新，下次操作时会重新计算
                                break;

                            case "delete":
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 删除操作：删除句子{startIndex}-{endIndex}");
                                SentenceState.ProcessSentenceDelete(startIndex, endIndex);
                                // 删除操作后，快照已更新，下次操作时会重新计算
                                break;

                            case "insert":
                                // 将insert内容切分为句子，为每个句子创建名字
                                List<string> insertSentences = SplitIntoSentences(action.Content ?? "");
                                if (insertSentences.Count == 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[快照操作] 警告：insert内容无法切分为句子，跳过操作");
                                    continue;
                                }
                                
                                System.Diagnostics.Debug.WriteLine($"[快照操作] 插入操作：在句子{startIndex}的{action.Position}位置插入{insertSentences.Count}个新句子");
                                SentenceState.ProcessSentenceInsert(startIndex, action.Position, insertSentences);
                                // 插入操作后，快照已更新，下次操作时会重新计算
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
                // 获取句子名称（去掉+号前缀，如果有的话）
                string displayHunkSentenceName = displayHunkLine.StartsWith("+")
                    ? displayHunkLine.Substring(1)
                    : displayHunkLine;
                
                string snapshotSentenceName = currentSnapshot[snapshotIndex];
                
                // 记录映射：display_hunk位置（从1开始）-> 快照位置（从1开始）
                int displayHunkPosition = displayHunkIndex + 1;
                int snapshotPosition = snapshotIndex + 1;
                mapping[displayHunkPosition] = snapshotPosition;
                
                System.Diagnostics.Debug.WriteLine($"[映射构建] display_hunk位置{displayHunkPosition}({displayHunkSentenceName}) -> 快照位置{snapshotPosition}({snapshotSentenceName})");
                
                // 两个指针都移动
                displayHunkIndex++;
                snapshotIndex++;
            }
            
            return mapping;
        }

        /// <summary>
        /// 获取display_hunk中的句子总数
        /// </summary>
        private static int GetDisplayHunkSentenceCount(IReadOnlyList<(string header, IReadOnlyList<string> lines)> displayHunk)
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
        /// 在段落中查找句子的精确Range位置
        /// </summary>
        /// <param name="paragraph">段落对象</param>
        /// <param name="sentenceText">要查找的句子文本</param>
        /// <returns>句子的Range，如果找不到则返回null</returns>
        private static Word.Range FindSentenceRangeInParagraph(Word.Paragraph paragraph, string sentenceText)
        {
            try
            {
                if (paragraph == null || string.IsNullOrEmpty(sentenceText))
                    return null;

                string paraText = paragraph.Range.Text?.TrimEnd('\r', '\n') ?? "";
                string originalParaText = paraText;
                
                // 找到按钮书签的位置（如果存在）
                int buttonStartInPara = -1;
                int buttonEndInPara = -1;
                Word.Document doc = paragraph.Range.Document;
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    if (bookmark.Name.StartsWith("operation_"))
                    {
                        Word.Range bookmarkRange = bookmark.Range;
                        // 检查书签是否在段落范围内
                        if (bookmarkRange.Start >= paragraph.Range.Start && bookmarkRange.End <= paragraph.Range.End)
                        {
                            // 获取书签在段落中的相对位置
                            buttonStartInPara = bookmarkRange.Start - paragraph.Range.Start;
                            buttonEndInPara = bookmarkRange.End - paragraph.Range.Start;
                            break;
                        }
                    }
                }
                
                // 去除按钮文本后切分句子
                paraText = RemoveButtonTextFromContent(doc, paragraph.Range, paraText);
                
                // 使用更精确的匹配：先切分句子，然后匹配
                List<SentenceInfo> sentenceInfos = SplitIntoSentencesWithPosition(paraText);
                
                // 查找匹配的句子（使用Trim后的文本比较）
                string trimmedSentenceText = sentenceText.Trim();
                foreach (var sentenceInfo in sentenceInfos)
                {
                    if (sentenceInfo.Text.Trim() == trimmedSentenceText)
                    {
                        // 找到匹配的句子，计算在Word文档中的位置
                        int paraRangeStart = paragraph.Range.Start;
                        
                        // 计算句子在去除按钮文本后的文本中的位置
                        int sentenceStartInCleanText = sentenceInfo.StartIndex;
                        int sentenceEndInCleanText = sentenceInfo.EndIndex + 1; // +1因为EndIndex是包含的
                        
                        // 计算偏移量：需要将去除按钮文本后的位置转换回原始位置
                        int offset = 0;
                        if (buttonStartInPara >= 0)
                        {
                            int buttonTextLength = buttonEndInPara - buttonStartInPara;
                            // RemoveButtonTextFromContent从字符串中删除了按钮文本
                            // 如果句子在按钮之后，去除按钮后句子的位置会减少buttonTextLength
                            // 所以需要加上buttonTextLength来恢复原始位置
                            if (sentenceStartInCleanText >= buttonStartInPara)
                            {
                                // 句子在按钮之后（或相同位置），需要加上按钮文本的长度
                                offset = buttonTextLength;
                            }
                            // 如果句子在按钮之前，不需要调整（offset = 0）
                        }
                        
                        // 计算句子在Word文档中的精确位置
                        int sentenceStart = paraRangeStart + sentenceStartInCleanText + offset;
                        int sentenceEnd = paraRangeStart + sentenceEndInCleanText + offset;
                        
                        // 创建Range
                        Word.Range sentenceRange = doc.Range(sentenceStart, sentenceEnd);
                        return sentenceRange;
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 查找句子Range失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 在整个文档中查找句子的Range位置（不依赖段落）
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="sentenceText">句子文本</param>
        /// <returns>句子的Range</returns>
        private static Word.Range FindSentenceRangeInDocument(Word.Document doc, string sentenceText)
        {
            try
            {
                if (doc == null || string.IsNullOrEmpty(sentenceText))
                    return null;

                // 获取整个文档内容（保留换行符），不去除按钮文本
                // 因为SentenceState中的句子内容不包含按钮文本，所以可以直接在文档中搜索
                string fullText = doc.Content.Text ?? "";
                
                // 使用更精确的匹配：先切分句子，然后匹配
                List<SentenceInfo> sentenceInfos = SplitIntoSentencesWithPosition(fullText);
                
                // 查找匹配的句子（只去除空格和制表符，保留换行符）
                string normalizedSentenceText = sentenceText.TrimStart(' ', '\t').TrimEnd(' ', '\t');
                foreach (var sentenceInfo in sentenceInfos)
                {
                    string normalizedInfoText = sentenceInfo.Text.TrimStart(' ', '\t').TrimEnd(' ', '\t');
                    if (normalizedInfoText == normalizedSentenceText)
                    {
                        // 找到匹配的句子，计算在Word文档中的位置
                        int docRangeStart = doc.Content.Start;
                        
                        // 句子在文档文本中的位置（不去除按钮，直接使用原始位置）
                        int sentenceStartInText = sentenceInfo.StartIndex;
                        int sentenceEndInText = sentenceInfo.EndIndex + 1; // +1因为EndIndex是包含的
                        
                        // 计算句子在Word文档中的精确位置
                        int sentenceStart = docRangeStart + sentenceStartInText;
                        int sentenceEnd = docRangeStart + sentenceEndInText;
                        
                        // 创建Range
                        Word.Range sentenceRange = doc.Range(sentenceStart, sentenceEnd);
                        
                        // 调试信息：打印查找过程
                        string sentenceDisplayText = sentenceText.Replace("\r", "\\r").Replace("\n", "\\n");
                        System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 查找句子: \"{sentenceDisplayText}\"");
                        System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 文档原始长度: {doc.Content.Text?.Length ?? 0}");
                        System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 在文档文本中的位置: Start={sentenceStartInText}, End={sentenceEndInText}");
                        System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 在Word文档中的位置: Start={sentenceStart}, End={sentenceEnd}");
                        string foundText = sentenceRange.Text ?? "";
                        string foundDisplayText = foundText.Replace("\r", "\\r").Replace("\n", "\\n");
                        System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 找到的Range内容: \"{foundDisplayText}\"");
                        
                        return sentenceRange;
                    }
                }
                
                string sentenceDisplayTextNotFound = sentenceText.Replace("\r", "\\r").Replace("\n", "\\n");
                System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 未找到匹配的句子: \"{sentenceDisplayTextNotFound}\"");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 查找句子Range失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 在段落中查找多个句子的Range位置
        /// </summary>
        /// <param name="paragraph">段落对象</param>
        /// <param name="sentences">句子文本列表</param>
        /// <returns>句子Range列表（按顺序）</returns>
        private static List<Word.Range> FindSentenceRangesInParagraph(Word.Paragraph paragraph, List<string> sentences)
        {
            var ranges = new List<Word.Range>();
            if (paragraph == null || sentences == null || sentences.Count == 0)
                return ranges;

            string paraText = paragraph.Range.Text?.TrimEnd('\r', '\n') ?? "";
            paraText = RemoveButtonTextFromContent(paragraph.Range.Document, paragraph.Range, paraText);
            
            int searchStart = 0;
            foreach (var sentence in sentences)
            {
                if (string.IsNullOrEmpty(sentence))
                    continue;

                int textIndex = paraText.IndexOf(sentence, searchStart, StringComparison.Ordinal);
                if (textIndex < 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 在段落中找不到句子: {sentence.Substring(0, Math.Min(30, sentence.Length))}");
                    break;
                }

                // 计算在Word文档中的位置
                int paraRangeStart = paragraph.Range.Start;
                string fullParaText = paragraph.Range.Text ?? "";
                int buttonTextLength = fullParaText.Length - paraText.Length - 1;
                if (buttonTextLength < 0) buttonTextLength = 0;
                
                int sentenceStart = paraRangeStart + textIndex + buttonTextLength;
                int sentenceEnd = sentenceStart + sentence.Length;
                
                Word.Range sentenceRange = paragraph.Range.Document.Range(sentenceStart, sentenceEnd);
                ranges.Add(sentenceRange);
                
                searchStart = textIndex + sentence.Length;
            }

            return ranges;
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
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] originalFullDisplayContent: {string.Join(",", originalFullDisplayContent)}");
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] targetFullDisplayContent: {string.Join(",", targetFullDisplayContent)}");

                // 第一步：使用LCS算法找到保持顺序一致的唯一锚点序列
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤2: 开始使用LCS算法找到唯一锚点序列");
                var anchorSequences = FindOrderedAnchorPoints(originalFullDisplayContent, targetFullDisplayContent);
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤2完成: 找到 {anchorSequences.Count} 个唯一锚点序列");

                // 第二步：删除相邻锚点之间的内容（直接操作句子，不依赖段落）
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤3: 开始删除相邻锚点之间的内容");
                var sentencesToDelete = new HashSet<string>(); // 使用HashSet避免重复

                // 遍历相邻的锚点序列对
                for (int i = 0; i < anchorSequences.Count - 1; i++)
                {
                    var startSequence = anchorSequences[i];
                    var endSequence = anchorSequences[i + 1];

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 处理锚点序列区间 [{string.Join(",", startSequence.sequence)}(original: {startSequence.originalStartIndex}-{startSequence.originalEndIndex}, target: {startSequence.targetStartIndex}-{startSequence.targetEndIndex}) - {string.Join(",", endSequence.sequence)}(original: {endSequence.originalStartIndex}-{endSequence.originalEndIndex}, target: {endSequence.targetStartIndex}-{endSequence.targetEndIndex})]");

                    // 收集原始序列中这个区间的所有句子（包括锚点）
                    var originalSegmentList = new List<string>();
                    // 先添加开始锚点
                    if (startSequence.sequence.Count > 0)
                    {
                        originalSegmentList.AddRange(startSequence.sequence);
                    }
                    // 添加区间内的句子
                    for (int j = startSequence.originalEndIndex + 1; j < endSequence.originalStartIndex; j++)
                    {
                        if (j >= 0 && j < originalFullDisplayContent.Count)
                        {
                            originalSegmentList.Add(originalFullDisplayContent[j]);
                        }
                    }
                    // 添加结束锚点
                    if (endSequence.sequence.Count > 0)
                    {
                        originalSegmentList.AddRange(endSequence.sequence);
                    }
                    var originalSegment = new HashSet<string>(originalSegmentList);

                    // 收集目标序列中这个区间的所有句子（包括锚点）
                    var targetSegmentList = new List<string>();
                    // 先添加开始锚点
                    if (startSequence.sequence.Count > 0)
                    {
                        targetSegmentList.AddRange(startSequence.sequence);
                    }
                    // 添加区间内的句子
                    for (int j = startSequence.targetEndIndex + 1; j < endSequence.targetStartIndex; j++)
                    {
                        if (j >= 0 && j < targetFullDisplayContent.Count)
                        {
                            targetSegmentList.Add(targetFullDisplayContent[j]);
                        }
                    }
                    // 添加结束锚点
                    if (endSequence.sequence.Count > 0)
                    {
                        targetSegmentList.AddRange(endSequence.sequence);
                    }
                    var targetSegment = new HashSet<string>(targetSegmentList);

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 原始序列: {string.Join(" ", originalSegmentList)}");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 目标序列: {string.Join(" ", targetSegmentList)}");

                    // 找出在原始序列区间内但不在目标序列区间内的句子，标记为删除（不包括锚点）
                    var originalIntervalSentences = new HashSet<string>();
                    for (int j = startSequence.originalEndIndex + 1; j < endSequence.originalStartIndex; j++)
                    {
                        if (j >= 0 && j < originalFullDisplayContent.Count)
                        {
                            originalIntervalSentences.Add(originalFullDisplayContent[j]);
                        }
                    }
                    
                    var targetIntervalSentences = new HashSet<string>();
                    for (int j = startSequence.targetEndIndex + 1; j < endSequence.targetStartIndex; j++)
                    {
                        if (j >= 0 && j < targetFullDisplayContent.Count)
                        {
                            targetIntervalSentences.Add(targetFullDisplayContent[j]);
                        }
                    }
                    
                    var intervalSentencesToDelete = new List<string>();
                    foreach (var sentenceName in originalIntervalSentences)
                    {
                        if (!targetIntervalSentences.Contains(sentenceName))
                        {
                            intervalSentencesToDelete.Add(sentenceName);
                            sentencesToDelete.Add(sentenceName); // 使用HashSet自动去重
                        }
                    }
                    
                    if (intervalSentencesToDelete.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 区间 [{string.Join(",", startSequence.sequence)} - {string.Join(",", endSequence.sequence)}] 标记删除: {string.Join(",", intervalSentencesToDelete)}");
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤3完成: 共标记 {sentencesToDelete.Count} 个句子需要删除");
                
                // 处理最后一个锚点序列之后的内容（直接操作句子，不依赖段落）
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤4: 处理最后一个锚点序列之后的内容");
                if (anchorSequences.Count > 0)
                {
                    var lastSequence = anchorSequences[anchorSequences.Count - 1];

                    // 收集原始序列中最后一个锚点之后的所有句子（包括锚点）
                    var originalAfterLastAnchorList = new List<string>();
                    // 先添加最后一个锚点序列
                    if (lastSequence.sequence.Count > 0)
                    {
                        originalAfterLastAnchorList.AddRange(lastSequence.sequence);
                    }
                    // 添加之后的内容
                    for (int j = lastSequence.originalEndIndex + 1; j < originalFullDisplayContent.Count; j++)
                    {
                        if (j >= 0 && j < originalFullDisplayContent.Count)
                        {
                            originalAfterLastAnchorList.Add(originalFullDisplayContent[j]);
                        }
                    }
                    var originalAfterLastAnchor = new HashSet<string>(originalAfterLastAnchorList);

                    // 收集目标序列中最后一个锚点之后的所有句子（包括锚点）
                    var targetAfterLastAnchorList = new List<string>();
                    // 先添加最后一个锚点序列
                    if (lastSequence.sequence.Count > 0)
                    {
                        targetAfterLastAnchorList.AddRange(lastSequence.sequence);
                    }
                    // 添加之后的内容
                    for (int j = lastSequence.targetEndIndex + 1; j < targetFullDisplayContent.Count; j++)
                    {
                        if (j >= 0 && j < targetFullDisplayContent.Count)
                        {
                            targetAfterLastAnchorList.Add(targetFullDisplayContent[j]);
                        }
                    }
                    var targetAfterLastAnchor = new HashSet<string>(targetAfterLastAnchorList);

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 原始序列: {string.Join(" ", originalAfterLastAnchorList)}");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 目标序列: {string.Join(" ", targetAfterLastAnchorList)}");

                    // 找出在原始序列最后一个锚点之后但不在目标序列最后一个锚点之后的句子，标记为删除（不包括锚点）
                    var originalAfterLastAnchorOnly = new HashSet<string>();
                    for (int j = lastSequence.originalEndIndex + 1; j < originalFullDisplayContent.Count; j++)
                    {
                        if (j >= 0 && j < originalFullDisplayContent.Count)
                        {
                            originalAfterLastAnchorOnly.Add(originalFullDisplayContent[j]);
                        }
                    }
                    
                    var targetAfterLastAnchorOnly = new HashSet<string>();
                    for (int j = lastSequence.targetEndIndex + 1; j < targetFullDisplayContent.Count; j++)
                    {
                        if (j >= 0 && j < targetFullDisplayContent.Count)
                        {
                            targetAfterLastAnchorOnly.Add(targetFullDisplayContent[j]);
                        }
                    }
                    
                    var sentencesToDeleteAfter = new List<string>();
                    foreach (var sentenceName in originalAfterLastAnchorOnly)
                    {
                        if (!targetAfterLastAnchorOnly.Contains(sentenceName))
                        {
                            sentencesToDeleteAfter.Add(sentenceName);
                            sentencesToDelete.Add(sentenceName); // 添加到总删除列表
                        }
                    }
                    
                    if (sentencesToDeleteAfter.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点序列之后标记删除: {string.Join(",", sentencesToDeleteAfter)}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点序列之后无需删除");
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤4完成: 标记删除完成，共标记 {sentencesToDelete.Count} 个句子需要删除");
                
                // 统一执行所有删除操作（从后往前删除，避免位置偏移）
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤5: 执行删除操作");
                var sentencesToDeleteWithPosition = new List<(string name, int position)>();
                foreach (var sentenceName in sentencesToDelete)
                {
                    // 从名称句子内容映射表获取内容，然后查找range
                    string sentenceContent = SentenceState.GetSentenceContent(sentenceName);
                    if (!string.IsNullOrEmpty(sentenceContent))
                    {
                        Word.Range sentenceRange = FindSentenceRangeInDocument(doc, sentenceContent);
                        if (sentenceRange != null)
                        {
                            sentencesToDeleteWithPosition.Add((sentenceName, sentenceRange.Start));
                        }
                    }
                }
                
                foreach (var (name, pos) in sentencesToDeleteWithPosition.OrderByDescending(x => x.position))
                {
                    // 从名称句子内容映射表获取内容，然后查找range
                    string sentenceContent = SentenceState.GetSentenceContent(name);
                    if (!string.IsNullOrEmpty(sentenceContent))
                    {
                        Word.Range sentenceRange = FindSentenceRangeInDocument(doc, sentenceContent);
                        if (sentenceRange != null)
                        {
                            sentenceRange.Delete();
                            System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 已删除句子: {name}");
                        }
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤5完成: 共删除 {sentencesToDelete.Count} 个句子");
                
                // 按照target的顺序，在相邻锚点序列之间插入缺失的句子
                // 从后往前遍历锚点区间，避免前面插入影响后面位置
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤7: 按照target的顺序插入缺失的句子（从后往前处理）");
                for (int intervalIndex = anchorSequences.Count - 2; intervalIndex >= 0; intervalIndex--)
                {
                    var startSequence = anchorSequences[intervalIndex];
                    var endSequence = anchorSequences[intervalIndex + 1];

                    // 收集target中这个区间应该有的句子（按顺序）
                    var targetSegment = new List<string>();
                    for (int j = startSequence.targetEndIndex + 1; j < endSequence.targetStartIndex; j++)
                    {
                        if (j < targetFullDisplayContent.Count)
                        {
                            targetSegment.Add(targetFullDisplayContent[j]);
                        }
                    }

                    // 收集删除后的原始序列中这个区间的句子（从原始序列中获取）
                    // 由于我们不再维护映射表，直接从原始序列中获取区间内的句子
                    var originalSegment = new List<string>();
                    for (int j = startSequence.originalEndIndex + 1; j < endSequence.originalStartIndex; j++)
                    {
                        if (j >= 0 && j < originalFullDisplayContent.Count)
                        {
                            originalSegment.Add(originalFullDisplayContent[j]);
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 区间 [{string.Join(",", startSequence.sequence)} - {string.Join(",", endSequence.sequence)}]");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 原始序列: {string.Join(",", originalSegment)}");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 目标序列: {string.Join(",", targetSegment)}");

                    // 判断需要插入的句子：在targetSegment中但不在originalSegment中的句子
                    var sentencesToInsert = targetSegment.Where(s => !originalSegment.Contains(s)).ToList();

                    if (sentencesToInsert.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 该区间无需插入");
                        continue;
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 需要插入的句子: {string.Join(",", sentencesToInsert)}");

                    // 找到开始锚点序列的最后一个句子的Range（作为插入位置）
                    Word.Range anchorRange = null;
                    if (startSequence.sequence.Count > 0)
                    {
                        var lastAnchorName = startSequence.sequence.Last();
                        string anchorContent = SentenceState.GetSentenceContent(lastAnchorName);
                        if (!string.IsNullOrEmpty(anchorContent))
                        {
                            anchorRange = FindSentenceRangeInDocument(doc, anchorContent);
                        }
                    }

                    if (anchorRange == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 警告：找不到锚点Range，跳过该区间");
                        continue;
                    }

                    // 检查锚点句子内容是否包含换行符
                    bool anchorHasNewline = false;
                    if (startSequence.sequence.Count > 0)
                    {
                        var lastAnchorName = startSequence.sequence.Last();
                        string lastAnchorContent = SentenceState.GetSentenceContent(lastAnchorName);
                        anchorHasNewline = lastAnchorContent.Contains("\r\n") || lastAnchorContent.Contains("\r") || lastAnchorContent.Contains("\n");
                    }

                    // 将需要插入的所有句子作为一个整体拼接（直接按顺序拼接，句子内容已包含换行符）
                    var insertContentBuilder = new StringBuilder();
                    foreach (var sentenceName in sentencesToInsert)
                    {
                        string content = SentenceState.GetSentenceContent(sentenceName);
                        if (!string.IsNullOrEmpty(content))
                        {
                            insertContentBuilder.Append(content);
                        }
                    }
                    string insertContent = insertContentBuilder.ToString();

                    if (string.IsNullOrEmpty(insertContent))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 警告：插入内容为空，跳过");
                        continue;
                    }

                    // 在锚点序列后面插入
                    Word.Range insertRange = anchorRange;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    // 根据锚点是否有换行符来决定插入内容的格式
                    // 如果锚点有换行符，插入内容最后添加换行符（创建新段落）
                    // 如果锚点没有换行符，插入内容不添加换行符（追加到同一段落）
                    bool addedNewline = false;
                    if (anchorHasNewline)
                    {
                        // 锚点有换行符，插入内容最后添加换行符
                        // 先移除插入内容末尾可能已有的换行符，然后统一添加 \r
                        string trimmedContent = insertContent.TrimEnd('\r', '\n');
                        insertRange.Text = trimmedContent + "\r";
                        addedNewline = true;
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 锚点有换行符，插入内容添加换行符: {string.Join(",", sentencesToInsert)}");
                    }
                    else
                    {
                        // 锚点没有换行符，直接插入（追加到同一段落）
                        insertRange.Text = insertContent;
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 锚点无换行符，追加到同一段落: {string.Join(",", sentencesToInsert)}");
                    }

                    // 如果添加了换行符，需要更新最后一个句子的内容到名字句子内容映射表
                    if (addedNewline && sentencesToInsert.Count > 0)
                    {
                        var lastSentenceName = sentencesToInsert[sentencesToInsert.Count - 1];
                        string lastContent = SentenceState.GetSentenceContent(lastSentenceName);
                        if (!string.IsNullOrEmpty(lastContent))
                        {
                            // 移除末尾可能已有的换行符，然后添加 \r
                            string updatedContent = lastContent.TrimEnd('\r', '\n') + "\r";
                            SentenceState.UpdateSentenceContent(lastSentenceName, updatedContent);
                            System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 更新映射表: {lastSentenceName} 的内容添加换行符");
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤7完成: 锚点间插入完成，当前文档段落总数: {doc.Paragraphs.Count}");
                
                // 处理最后一个锚点序列之后的内容
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤8: 处理最后一个锚点序列之后的内容");
                if (anchorSequences.Count > 0)
                {
                    var lastSequence = anchorSequences[anchorSequences.Count - 1];

                    // 收集target中最后一个锚点序列之后应该有的句子
                    var targetAfterLastAnchor = new List<string>();
                    for (int j = lastSequence.targetEndIndex + 1; j < targetFullDisplayContent.Count; j++)
                    {
                        targetAfterLastAnchor.Add(targetFullDisplayContent[j]);
                    }

                    // 收集删除后的原始序列中最后一个锚点之后的句子（从原始序列中获取）
                    var originalAfterLastAnchor = new List<string>();
                    for (int j = lastSequence.originalEndIndex + 1; j < originalFullDisplayContent.Count; j++)
                    {
                        if (j >= 0 && j < originalFullDisplayContent.Count)
                        {
                            originalAfterLastAnchor.Add(originalFullDisplayContent[j]);
                        }
                    }
                    
                    // 查找最后一个锚点的Range（用于确定插入位置）
                    Word.Range lastAnchorRange = null;
                    if (lastSequence.sequence.Count > 0)
                    {
                        var lastAnchorName = lastSequence.sequence.Last();
                        string anchorContent = SentenceState.GetSentenceContent(lastAnchorName);
                        if (!string.IsNullOrEmpty(anchorContent))
                        {
                            lastAnchorRange = FindSentenceRangeInDocument(doc, anchorContent);
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点序列 {string.Join(",", lastSequence.sequence)}");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 原始序列之后: {string.Join(",", originalAfterLastAnchor)}");
                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 目标序列之后: {string.Join(",", targetAfterLastAnchor)}");

                    // 判断需要插入的句子：在targetAfterLastAnchor中但不在originalAfterLastAnchor中的句子
                    var sentencesToInsert = targetAfterLastAnchor.Where(s => !originalAfterLastAnchor.Contains(s)).ToList();

                    if (sentencesToInsert.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 最后一个锚点之后无需插入");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 需要插入的句子: {string.Join(",", sentencesToInsert)}");

                        // 使用之前获取的 lastAnchorRange（作为插入位置）
                        Word.Range anchorRange = lastAnchorRange;
                        if (anchorRange == null)
                        {
                            // 如果找不到锚点Range，在文档末尾插入
                            anchorRange = doc.Range(doc.Content.End - 1, doc.Content.End);
                        }

                        // 检查锚点句子内容是否包含换行符
                        bool anchorHasNewline = false;
                        if (lastSequence.sequence.Count > 0 && lastAnchorRange != null)
                        {
                            var lastAnchorName = lastSequence.sequence.Last();
                            string lastAnchorContent = SentenceState.GetSentenceContent(lastAnchorName);
                            anchorHasNewline = lastAnchorContent.Contains("\r\n") || lastAnchorContent.Contains("\r") || lastAnchorContent.Contains("\n");
                        }

                        // 将需要插入的所有句子作为一个整体拼接（直接按顺序拼接，句子内容已包含换行符）
                        var insertContentBuilder = new StringBuilder();
                        foreach (var sentenceName in sentencesToInsert)
                        {
                            string content = SentenceState.GetSentenceContent(sentenceName);
                            if (!string.IsNullOrEmpty(content))
                            {
                                insertContentBuilder.Append(content);
                            }
                        }
                        string insertContent = insertContentBuilder.ToString();

                        if (!string.IsNullOrEmpty(insertContent))
                        {
                            // 在锚点序列后面插入
                            Word.Range insertRange = anchorRange;
                            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                            // 根据锚点是否有换行符来决定插入内容的格式
                            // 如果锚点有换行符，插入内容最后添加换行符（创建新段落）
                            // 如果锚点没有换行符，插入内容不添加换行符（追加到同一段落）
                            bool addedNewline = false;
                            if (anchorHasNewline)
                            {
                                // 锚点有换行符，插入内容最后添加换行符
                                // 先移除插入内容末尾可能已有的换行符，然后统一添加 \r
                                string trimmedContent = insertContent.TrimEnd('\r', '\n');
                                insertRange.Text = trimmedContent + "\r";
                                addedNewline = true;
                                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 锚点有换行符，插入内容添加换行符: {string.Join(",", sentencesToInsert)}");
                            }
                            else
                            {
                                // 锚点没有换行符，直接插入（追加到同一段落）
                                insertRange.Text = insertContent;
                                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 锚点无换行符，追加到同一段落: {string.Join(",", sentencesToInsert)}");
                            }

                            // 如果添加了换行符，需要更新最后一个句子的内容到名字句子内容映射表
                            if (addedNewline && sentencesToInsert.Count > 0)
                            {
                                var lastSentenceName = sentencesToInsert[sentencesToInsert.Count - 1];
                                string lastContent = SentenceState.GetSentenceContent(lastSentenceName);
                                if (!string.IsNullOrEmpty(lastContent))
                                {
                                    // 移除末尾可能已有的换行符，然后添加 \r
                                    string updatedContent = lastContent.TrimEnd('\r', '\n') + "\r";
                                    SentenceState.UpdateSentenceContent(lastSentenceName, updatedContent);
                                    System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 更新映射表: {lastSentenceName} 的内容添加换行符");
                                }
                            }
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyHunkInsertions] 步骤8完成: 所有插入操作完成");

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

                // 第4-2步：应用格式和添加按钮（直接通过句子名称在整个文档中查找Range，不依赖段落索引）
                // 按顺序逐个处理操作块，确保每个操作块在添加按钮后，下一个操作块能基于最新的文档状态查找Range
                System.Diagnostics.Debug.WriteLine("[第四步-2] 应用格式和添加按钮（逐个操作块处理）");
                foreach (var block in operationBlocks)
                {
                    // 先查找Range并应用格式，然后添加按钮
                    // 这样下一个操作块查找Range时，文档已经包含了当前操作块的按钮
                    ApplyOperationBlockFormatting(doc, block, targetFullDisplayContent);
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
        /// 应用操作块格式和添加按钮（直接通过句子名称在整个文档中查找Range，不依赖段落索引）
        /// 处理顺序：先应用格式，再添加按钮，确保下一个操作块查找Range时基于最新的文档状态
        /// </summary>
        private static void ApplyOperationBlockFormatting(Word.Document doc, OperationBlock block, List<string> targetFullDisplayContent)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[操作块] 类型: {block.Type}, 行数: {block.Lines.Count}");

                // 第一步：先应用格式（在添加按钮之前，确保Range查找基于当前文档状态）
                System.Diagnostics.Debug.WriteLine($"[操作块格式] 开始应用格式");
                System.Diagnostics.Debug.WriteLine($"[操作块格式] 当前文档总长度: {doc.Content.Text?.Length ?? 0}");
                foreach (var line in block.Lines)
                {
                    if (line.StartsWith("-"))
                    {
                        // 负号的句子加上删除线
                        string sentenceName = line.Substring(1);
                        string sentenceContent = SentenceState.GetSentenceContent(sentenceName);
                        
                        if (!string.IsNullOrEmpty(sentenceContent))
                        {
                            // 直接在整个文档中查找句子Range
                            Word.Range sentenceRange = FindSentenceRangeInDocument(doc, sentenceContent);
                            if (sentenceRange != null)
                            {
                                // 应用格式到句子Range
                                ApplyDeleteFormatting(sentenceRange);
                                string rangeText = sentenceRange.Text ?? "";
                                string displayText = rangeText.Replace("\r", "\\r").Replace("\n", "\\n");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 句子{sentenceName}添加删除线，Range位置: Start={sentenceRange.Start}, End={sentenceRange.End}, 内容长度={sentenceRange.Text?.Length ?? 0}");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 句子{sentenceName}的Range内容: \"{displayText}\"");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 期望内容: \"{sentenceContent.Replace("\r", "\\r").Replace("\n", "\\n")}\"");
                            }
                            else
                            {
                                string contentDisplayText = sentenceContent.Replace("\r", "\\r").Replace("\n", "\\n");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 警告：找不到句子{sentenceName}的Range，内容: \"{contentDisplayText}\"");
                            }
                        }
                    }
                    else if (line.StartsWith("+"))
                    {
                        // 正号的句子加上黄色背景
                        string sentenceName = line.Substring(1);
                        string sentenceContent = SentenceState.GetSentenceContent(sentenceName);
                        
                        if (!string.IsNullOrEmpty(sentenceContent))
                        {
                            // 直接在整个文档中查找句子Range
                            Word.Range sentenceRange = FindSentenceRangeInDocument(doc, sentenceContent);
                            if (sentenceRange != null)
                            {
                                // 应用格式到句子Range
                                ApplyInsertFormatting(sentenceRange);
                                string rangeText = sentenceRange.Text ?? "";
                                string displayText = rangeText.Replace("\r", "\\r").Replace("\n", "\\n");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 句子{sentenceName}添加黄色背景，Range位置: Start={sentenceRange.Start}, End={sentenceRange.End}, 内容长度={sentenceRange.Text?.Length ?? 0}");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 句子{sentenceName}的Range内容: \"{displayText}\"");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 期望内容: \"{sentenceContent.Replace("\r", "\\r").Replace("\n", "\\n")}\"");
                            }
                            else
                            {
                                string contentDisplayText = sentenceContent.Replace("\r", "\\r").Replace("\n", "\\n");
                                System.Diagnostics.Debug.WriteLine($"[操作块格式] 警告：找不到句子{sentenceName}的Range，内容: \"{contentDisplayText}\"");
                            }
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[操作块格式] 格式应用完成，当前文档总长度: {doc.Content.Text?.Length ?? 0}");

                // 第二步：收集操作块中涉及的句子Range（用于按钮点击处理）
                Word.Range originalRange = null;
                Word.Range newRange = null;

                // 收集操作块中涉及的句子Range（直接在整个文档中搜索，不依赖段落）
                var deleteSentenceRanges = new List<Word.Range>();
                var insertSentenceRanges = new List<Word.Range>();

                foreach (var line in block.Lines)
                {
                    if (line.StartsWith("-"))
                    {
                        // 负号的句子
                        string sentenceName = line.Substring(1);
                        string sentenceContent = SentenceState.GetSentenceContent(sentenceName);
                        if (!string.IsNullOrEmpty(sentenceContent))
                        {
                            Word.Range sentenceRange = FindSentenceRangeInDocument(doc, sentenceContent);
                            if (sentenceRange != null)
                            {
                                deleteSentenceRanges.Add(sentenceRange);
                            }
                        }
                    }
                    else if (line.StartsWith("+"))
                    {
                        // 正号的句子
                        string sentenceName = line.Substring(1);
                        string sentenceContent = SentenceState.GetSentenceContent(sentenceName);
                        if (!string.IsNullOrEmpty(sentenceContent))
                        {
                            Word.Range sentenceRange = FindSentenceRangeInDocument(doc, sentenceContent);
                            if (sentenceRange != null)
                            {
                                insertSentenceRanges.Add(sentenceRange);
                            }
                        }
                    }
                }

                // 合并多个句子Range（用于按钮点击处理）
                if (deleteSentenceRanges.Count > 1)
                {
                    originalRange = doc.Range(deleteSentenceRanges[0].Start, deleteSentenceRanges[deleteSentenceRanges.Count - 1].End);
                }
                else if (deleteSentenceRanges.Count == 1)
                {
                    originalRange = deleteSentenceRanges[0];
                }

                if (insertSentenceRanges.Count > 1)
                {
                    newRange = doc.Range(insertSentenceRanges[0].Start, insertSentenceRanges[insertSentenceRanges.Count - 1].End);
                }
                else if (insertSentenceRanges.Count == 1)
                {
                    newRange = insertSentenceRanges[0];
                }

                // 确定按钮位置：根据操作类型找到对应的最后一个句子
                Word.Range buttonRange = null;
                int buttonStartPos = -1;
                
                // 根据操作类型确定按钮位置（直接在整个文档中搜索，不依赖段落）
                if (block.Type == "replace")
                {
                    // 替换操作：按钮放在新增的句子后面（最后一个+号句子）
                    string lastInsertSentenceName = null;
                    for (int i = block.Lines.Count - 1; i >= 0; i--)
                    {
                        if (block.Lines[i].StartsWith("+"))
                        {
                            lastInsertSentenceName = block.Lines[i].Substring(1);
                            break;
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(lastInsertSentenceName))
                    {
                        string sentenceContent = SentenceState.GetSentenceContent(lastInsertSentenceName);
                        if (!string.IsNullOrEmpty(sentenceContent))
                        {
                            buttonRange = FindSentenceRangeInDocument(doc, sentenceContent);
                            if (buttonRange != null)
                            {
                                string rangeText = buttonRange.Text ?? "";
                                string displayText = rangeText.Replace("\r", "\\r").Replace("\n", "\\n");
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 查找按钮位置，句子{lastInsertSentenceName}的Range: Start={buttonRange.Start}, End={buttonRange.End}, 内容长度={buttonRange.Text?.Length ?? 0}");
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 句子{lastInsertSentenceName}的Range内容: \"{displayText}\"");
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 期望内容: \"{sentenceContent.Replace("\r", "\\r").Replace("\n", "\\n")}\"");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 警告：找不到按钮位置句子{lastInsertSentenceName}的Range");
                            }
                        }
                    }
                }
                else if (block.Type == "delete")
                {
                    // 删除操作：按钮放在被删除的句子后面（最后一个-号句子）
                    string lastDeleteSentenceName = null;
                    for (int i = block.Lines.Count - 1; i >= 0; i--)
                    {
                        if (block.Lines[i].StartsWith("-"))
                        {
                            lastDeleteSentenceName = block.Lines[i].Substring(1);
                            break;
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(lastDeleteSentenceName))
                    {
                        string sentenceContent = SentenceState.GetSentenceContent(lastDeleteSentenceName);
                        if (!string.IsNullOrEmpty(sentenceContent))
                        {
                            buttonRange = FindSentenceRangeInDocument(doc, sentenceContent);
                        }
                    }
                }
                else if (block.Type == "insert")
                {
                    // 插入操作：按钮放在新增句子的后面（最后一个+号句子）
                    string lastInsertSentenceName = null;
                    for (int i = block.Lines.Count - 1; i >= 0; i--)
                    {
                        if (block.Lines[i].StartsWith("+"))
                        {
                            lastInsertSentenceName = block.Lines[i].Substring(1);
                            break;
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(lastInsertSentenceName))
                    {
                        string sentenceContent = SentenceState.GetSentenceContent(lastInsertSentenceName);
                        if (!string.IsNullOrEmpty(sentenceContent))
                        {
                            buttonRange = FindSentenceRangeInDocument(doc, sentenceContent);
                        }
                    }
                }

                // 收集操作块中的句子名称列表（用于按钮点击处理）
                var originalSentenceNames = new List<string>();
                var newSentenceNames = new List<string>();
                foreach (var line in block.Lines)
                {
                    if (line.StartsWith("-"))
                    {
                        originalSentenceNames.Add(line.Substring(1));
                    }
                    else if (line.StartsWith("+"))
                    {
                        newSentenceNames.Add(line.Substring(1));
                    }
                }

                // 插入按钮
                if (buttonRange != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[操作块按钮] 准备添加按钮，当前文档总长度: {doc.Content.Text?.Length ?? 0}");
                    try
                    {
                        // 检查句子内容是否以换行符结尾，如果是，按钮应该插入在句子末尾和换行符之间
                        string buttonSentenceName = null;
                        if (block.Type == "replace")
                        {
                            // 替换操作：使用最后一个+号句子
                            for (int i = block.Lines.Count - 1; i >= 0; i--)
                            {
                                if (block.Lines[i].StartsWith("+"))
                                {
                                    buttonSentenceName = block.Lines[i].Substring(1);
                                    break;
                                }
                            }
                        }
                        else if (block.Type == "delete")
                        {
                            // 删除操作：使用最后一个-号句子
                            for (int i = block.Lines.Count - 1; i >= 0; i--)
                            {
                                if (block.Lines[i].StartsWith("-"))
                                {
                                    buttonSentenceName = block.Lines[i].Substring(1);
                                    break;
                                }
                            }
                        }
                        else if (block.Type == "insert")
                        {
                            // 插入操作：使用最后一个+号句子
                            for (int i = block.Lines.Count - 1; i >= 0; i--)
                            {
                                if (block.Lines[i].StartsWith("+"))
                                {
                                    buttonSentenceName = block.Lines[i].Substring(1);
                                    break;
                                }
                            }
                        }
                        
                        // 如果句子内容以换行符结尾，调整按钮插入位置
                        if (!string.IsNullOrEmpty(buttonSentenceName))
                        {
                            string sentenceContent = SentenceState.GetSentenceContent(buttonSentenceName);
                            if (!string.IsNullOrEmpty(sentenceContent) && 
                                (sentenceContent.EndsWith("\r") || sentenceContent.EndsWith("\n") || sentenceContent.EndsWith("\r\n")))
                            {
                                // 创建一个新的Range，只包含句子内容（不包括换行符）
                                // 计算换行符的长度
                                int newlineLength = 0;
                                if (sentenceContent.EndsWith("\r\n"))
                                {
                                    newlineLength = 2;
                                }
                                else if (sentenceContent.EndsWith("\r") || sentenceContent.EndsWith("\n"))
                                {
                                    newlineLength = 1;
                                }
                                
                                // 创建不包含换行符的Range
                                Word.Range buttonInsertRange = doc.Range(buttonRange.Start, buttonRange.End - newlineLength);
                                buttonStartPos = buttonInsertRange.End;
                                
                                // 创建DocumentProcessor实例来添加按钮
                                var documentProcessor = new DocumentProcessor(doc.Application);
                                // 传递句子名称列表，以便按钮点击时直接使用
                                documentProcessor.InsertInlineButtons(buttonInsertRange, block.Type, originalRange, newRange, originalSentenceNames, newSentenceNames);
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] {block.Type}操作块按钮添加完成（在换行符前），按钮位置: {buttonInsertRange.Start}-{buttonInsertRange.End}");
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 原始句子名称: {string.Join(",", originalSentenceNames)}, 新句子名称: {string.Join(",", newSentenceNames)}");
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 添加按钮后文档总长度: {doc.Content.Text?.Length ?? 0}");
                            }
                            else
                            {
                                // 句子内容不以换行符结尾，正常插入
                                buttonStartPos = buttonRange.End;
                                
                                // 创建DocumentProcessor实例来添加按钮
                                var documentProcessor = new DocumentProcessor(doc.Application);
                                // 传递句子名称列表，以便按钮点击时直接使用
                                documentProcessor.InsertInlineButtons(buttonRange, block.Type, originalRange, newRange, originalSentenceNames, newSentenceNames);
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] {block.Type}操作块按钮添加完成，按钮位置: {buttonRange.Start}-{buttonRange.End}");
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 原始句子名称: {string.Join(",", originalSentenceNames)}, 新句子名称: {string.Join(",", newSentenceNames)}");
                                System.Diagnostics.Debug.WriteLine($"[操作块按钮] 添加按钮后文档总长度: {doc.Content.Text?.Length ?? 0}");
                            }
                        }
                        else
                        {
                            // 无法确定句子名称，使用默认方式
                            buttonStartPos = buttonRange.End;
                            
                            // 创建DocumentProcessor实例来添加按钮
                            var documentProcessor = new DocumentProcessor(doc.Application);
                            // 传递句子名称列表，以便按钮点击时直接使用
                            documentProcessor.InsertInlineButtons(buttonRange, block.Type, originalRange, newRange, originalSentenceNames, newSentenceNames);
                            System.Diagnostics.Debug.WriteLine($"[操作块按钮] {block.Type}操作块按钮添加完成，按钮位置: {buttonRange.Start}-{buttonRange.End}");
                            System.Diagnostics.Debug.WriteLine($"[操作块按钮] 原始句子名称: {string.Join(",", originalSentenceNames)}, 新句子名称: {string.Join(",", newSentenceNames)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 添加操作块按钮失败: {ex.Message}");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[操作块按钮] 警告：无法找到{block.Type}操作块的按钮位置");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用操作块格式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 验证句子序列在快照中是否连续
        /// </summary>
        /// <param name="sentenceNames">句子名称序列（逗号分隔）</param>
        /// <param name="actionType">操作类型（用于错误消息）</param>
        /// <returns>如果连续返回null，如果不连续返回错误消息</returns>
        private static string ValidateSentenceSequenceContinuity(string sentenceNames, string actionType)
        {
            if (string.IsNullOrEmpty(sentenceNames))
                return null;

            // 解析句子名称序列
            var sentenceList = sentenceNames.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            if (sentenceList.Count == 0)
                return null;

            // 获取当前快照
            var snapshot = SentenceState.CurrentSnapshot;
            if (snapshot == null || snapshot.Count == 0)
            {
                return $"错误：无法验证句子序列连续性，因为快照为空。请先调用F_get_document_content工具获取文档内容。";
            }

            // 在快照中查找每个句子的位置
            var positions = new List<(string name, int index)>();
            var snapshotList = snapshot.ToList(); // 转换为List以便使用IndexOf
            foreach (var sentenceName in sentenceList)
            {
                int index = snapshotList.IndexOf(sentenceName);
                if (index < 0)
                {
                    return $"错误：句子 \"{sentenceName}\" 在快照中不存在。请先调用F_get_document_content工具获取当前文档的句子编号序列。";
                }
                positions.Add((sentenceName, index));
            }

            // 按位置排序
            positions = positions.OrderBy(p => p.index).ToList();

            // 检查是否连续
            var groups = new List<List<string>>();
            var currentGroup = new List<string> { positions[0].name };

            for (int i = 1; i < positions.Count; i++)
            {
                // 如果当前位置和前一个位置连续（相差1），添加到当前组
                if (positions[i].index == positions[i - 1].index + 1)
                {
                    currentGroup.Add(positions[i].name);
                }
                else
                {
                    // 不连续，开始新组
                    groups.Add(currentGroup);
                    currentGroup = new List<string> { positions[i].name };
                }
            }
            groups.Add(currentGroup);

            // 如果只有一个组，说明连续
            if (groups.Count == 1)
                return null;

            // 不连续，构造错误消息
            // 格式："AAAAA","AAABD" 和 "AAAQA","ABAQA"
            var groupStrings = groups.Select(g => "\"" + string.Join("\",\"", g) + "\"").ToList();
            var groupDetails = string.Join(" 和 ", groupStrings);

            // 构造建议的正确格式
            var suggestedActions = new List<string>();
            for (int i = 0; i < groups.Count; i++)
            {
                var groupSentenceNames = string.Join(",", groups[i]);
                if (actionType.ToLower().StartsWith("replace"))
                {
                    suggestedActions.Add($"{{\"type\": \"replace\",\"detail\": {{\"original\": {{\"sentence_names\": \"{groupSentenceNames}\"}},\"new\": \"新内容{i + 1}\"}}}}");
                }
                else if (actionType.ToLower().StartsWith("delete"))
                {
                    suggestedActions.Add($"{{\"type\": \"delete\",\"detail\": {{\"deletion\": {{\"sentence_names\": \"{groupSentenceNames}\"}}}}}}");
                }
                else if (actionType.ToLower().StartsWith("insert"))
                {
                    // insert操作需要pre或post，这里简化处理
                    suggestedActions.Add($"{{\"type\": \"insert\",\"detail\": {{\"pre\": {{\"sentence_names\": \"{groupSentenceNames}\"}},\"insert\": \"插入内容{i + 1}\"}}}}");
                }
            }

            var suggestedFormat = "{\"action\": [" + string.Join(",", suggestedActions) + "]}";

            return $"{groupDetails}不连续，需要分开来写，应该的结构是：{suggestedFormat}";
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
            public int TargetIndex { get; set; } // 句子序号（在ProcessActionsForSnapshot中通过句子编号序列查找后设置）
            public string Position { get; set; }
            public string Content { get; set; }
            
            // 用于通过句子编号定位的字段
            public string SentenceNames { get; set; } // 句子编号序列，如 "AAAAA,AAAAB,AAAAC"（replace和delete操作使用）
            public string PreSentenceNames { get; set; } // 上文句子编号序列（insert操作使用）
            public string PostSentenceNames { get; set; } // 下文句子编号序列（insert操作使用）
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



