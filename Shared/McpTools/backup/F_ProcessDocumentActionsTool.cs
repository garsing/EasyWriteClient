using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 文档编辑操作工具
    /// 批量处理Word文档编辑操作
    /// </summary>
    public static class F_ProcessDocumentActionsTool
    {
        /// <summary>
        /// 将位置列表转换为区间列表
        /// </summary>
        private static List<(int Start, int End)> ConvertPositionsToRanges(List<int> positions)
        {
            var ranges = new List<(int Start, int End)>();
            if (positions == null || positions.Count == 0)
            {
                return ranges;
            }

            int start = positions[0];
            int current = positions[0];

            for (int i = 1; i < positions.Count; i++)
            {
                if (positions[i] == current + 1)
                {
                    // 连续，继续扩展
                    current = positions[i];
                }
                else
                {
                    // 不连续，保存当前区间，开始新区间
                    ranges.Add((start, current));
                    start = positions[i];
                    current = positions[i];
                }
            }

            // 添加最后一个区间
            ranges.Add((start, current));

            return ranges;
        }

        /// <summary>
        /// 第9步：根据历史记录进行格式和按钮操作
        /// </summary>
        private static void ApplyFormattingAndButtons(DocumentProcessor documentProcessor, int operationIndex)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{operationIndex}次新修改-第9步：根据历史记录进行格式和按钮操作");

            // 重新生成操作块（基于当前历史记录）
            var (updatedHistoricalDeletes, updatedHistoricalInserts) = DocumentState.GetConsolidatedHistoricalRanges();
            var operationBlocks = GenerateOperationBlocks(updatedHistoricalDeletes, updatedHistoricalInserts);

            // 应用格式和按钮
            ApplyOperationBlocksToDocument(operationBlocks, documentProcessor);
        }

        /// <summary>
        /// 注册文档编辑操作工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            List<McpTool> availableTools,
            object wordApplication)
        {
            var tool = new McpTool
            {
                name = "F_process_document_actions",
                alias = "处理文档操作",
                description = @"对Word文档执行插入、替换或删除操作。支持三种插入方式：传统定位插入、直接位置插入、光标位置插入。传统方式通过start/end字符精确定位文本，避免输入完整文本出错。

⚠️关键提醒：start和end字符本身都包含在要操作的文本范围内！

重要提示：
- start：只需提供原文开头的一小段特征文字（建议5-15个字符）
- end：只需提供原文结尾的一小段特征文字（建议5-15个字符）
- start和end应该是不同的文字，用来精确定位要操作的文本范围
- ⚠️系统会找到从start到end之间的所有文本进行操作（包含start和end字符）
- ⚠️start和end字符本身会包含在选中的文本范围内，不要遗漏end字符！

插入操作的三种方式：
1. 传统定位插入：使用pre/post上下文定位插入位置
2. 直接位置插入：使用position参数直接指定插入位置（0-based字符位置）
3. 光标位置插入：使用cursor=true在当前光标位置插入

示例输入格式：
{
  ""action"": [
    {
      ""type"": ""replace"",
      ""detail"": {
        ""original"": {
          ""start"": ""人工智能（AI）是"",
          ""end"": ""任务的系统""
        },
        ""new"": ""人工智能（AI）是计算机科学的重要分支，其核心目标是构建能够模拟人类智能行为的系统，能够执行通常需要人类智慧才能完成的任务""
      }
    },
    {
      ""type"": ""delete"",
      ""detail"": {
        ""deletion"": {
          ""start"": ""的发展历程"",
          ""end"": ""概念可以追溯""
        }
      }
    },
    {
      ""type"": ""insert"",
      ""detail"": {
        ""pre"": ""上文内容"",
        ""insert"": ""要插入的新内容"",
        ""post"": ""下文内容""
      }
    }
  ]
}

⚠️重要强调：start和end字符本身都包含在选中的文本范围内！",
                usage = @"对Word文档执行插入、替换或删除操作。支持三种插入方式：传统定位插入、直接位置插入、光标位置插入。传统方式通过start/end字符精确定位文本，避免输入完整文本出错。

⚠️关键提醒：start和end字符本身都包含在要操作的文本范围内！

重要提示：
- start：只需提供原文开头的一小段特征文字（建议5-15个字符）
- end：只需提供原文结尾的一小段特征文字（建议5-15个字符）
- start和end应该是不同的文字，用来精确定位要操作的文本范围
- ⚠️系统会找到从start到end之间的所有文本进行操作（包含start和end字符）
- ⚠️start和end字符本身会包含在选中的文本范围内，不要遗漏end字符！

插入操作的三种方式：
1. 传统定位插入：使用pre/post上下文定位插入位置
2. 直接位置插入：使用position参数直接指定插入位置（0-based字符位置）
3. 光标位置插入：使用cursor=true在当前光标位置插入

示例输入格式：
{
  ""action"": [
    {
      ""type"": ""replace"",
      ""detail"": {
        ""original"": {
          ""start"": ""人工智能（AI）是"",
          ""end"": ""任务的系统""
        },
        ""new"": ""人工智能（AI）是计算机科学的重要分支，其核心目标是构建能够模拟人类智能行为的系统，能够执行通常需要人类智慧才能完成的任务""
      }
    },
    {
      ""type"": ""delete"",
      ""detail"": {
        ""deletion"": {
          ""start"": ""的发展历程"",
          ""end"": ""概念可以追溯""
        }
      }
    },
    {
      ""type"": ""insert"",
      ""detail"": {
        ""pre"": ""上文内容"",
        ""insert"": ""要插入的新内容"",
        ""post"": ""下文内容""
      }
    }
  ]
}

⚠️重要强调：start和end字符本身都包含在选中的文本范围内！",
                input_schema = new McpTool.ToolInputSchema
                {
                    Type = "object",
                    Properties = new Dictionary<string, object>
                    {
                        ["action"] = new
                        {
                            type = "array",
                            description = "要执行的操作数组",
                            items = new
                            {
                                type = "object",
                                properties = new Dictionary<string, object>
                                {
                                    ["type"] = new
                                    {
                                        type = "string",
                                        description = "操作类型：'insert'(插入)、'replace'(替换)、'delete'(删除)",
                                        @enum = new[] { "insert", "replace", "delete" }
                                    },
                                    ["detail"] = new
                                    {
                                        type = "object",
                                        description = @"操作参数详情，根据操作类型提供相应参数。
⚠️重要提示：使用start/end格式时，只需提供少量特征文字即可精确定位，start和end字符本身都包含在选中的文本范围内！",
                                        properties = new Dictionary<string, object>
                                        {
                                            // insert操作的参数
                                            ["pre"] = new
                                            {
                                                type = "string",
                                                description = "insert操作：上文内容，用于定位插入位置（传统方式）"
                                            },
                                            ["insert"] = new
                                            {
                                                type = "string",
                                                description = "insert操作：要插入的新内容"
                                            },
                                            ["post"] = new
                                            {
                                                type = "string",
                                                description = "insert操作：下文内容，用于定位插入位置（传统方式）"
                                            },
                                            ["position"] = new
                                            {
                                                type = "integer",
                                                description = "insert操作：直接指定插入位置（从文档开头算起的字符位置，0-based）"
                                            },
                                            ["cursor"] = new
                                            {
                                                type = "boolean",
                                                description = "insert操作：是否在当前光标位置插入（true表示在光标位置插入）"
                                            },
                                            // replace操作的参数
                                            ["original"] = new
                                            {
                                                oneOf = new object[]
                                                {
                                                    new { type = "string", description = "replace操作：完整的原文内容（兼容旧格式）" },
                                                    new
                                                    {
                                                        type = "object",
                                                        description = "replace操作：通过start/end字符精确定位原文",
                                                        properties = new Dictionary<string, object>
                                                        {
                                                            ["start"] = new { type = "string", description = "原文开头的一小段特征文字（请尽量简短，建议5-15个字符，用于精确定位，避免输入整个原文内容）" },
                                                            ["end"] = new { type = "string", description = "原文结尾的一小段特征文字（请尽量简短，建议5-15个字符，用于精确定位，避免输入整个原文内容）。⚠️重要：end字符本身会包含在要操作的文本范围内！" }
                                                        },
                                                        required = new[] { "start", "end" }
                                                    }
                                                }
                                            },
                                            ["new"] = new
                                            {
                                                type = "string",
                                                description = "replace操作：新的文本内容"
                                            },
                                            // delete操作的参数
                                            ["deletion"] = new
                                            {
                                                oneOf = new object[]
                                                {
                                                    new { type = "string", description = "delete操作：完整的删除文本（兼容旧格式）" },
                                                    new
                                                    {
                                                        type = "object",
                                                        description = "delete操作：通过start/end字符精确定位要删除的文本",
                                                        properties = new Dictionary<string, object>
                                                        {
                                                            ["start"] = new { type = "string", description = "要删除文本开头的一小段特征文字（请尽量简短，建议5-15个字符，用于精确定位，避免输入整个删除内容）" },
                                                            ["end"] = new { type = "string", description = "要删除文本结尾的一小段特征文字（请尽量简短，建议5-15个字符，用于精确定位，避免输入整个删除内容）。⚠️重要：end字符本身会包含在要删除的文本范围内！" }
                                                        },
                                                        required = new[] { "start", "end" }
                                                    }
                                                }
                                            }
                                        },
                                        additionalProperties = true
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
                System.Diagnostics.Debug.WriteLine("[DEBUG] process_document_actions工具开始执行");
                try
                {
                    if (wordApplication == null)
                    {
                        System.Diagnostics.Debug.WriteLine("[DEBUG] Word应用程序实例不可用");
                        return new ToolResult { Success = false, Error = "Word应用程序实例不可用" };
                    }

                    // 创建DocumentProcessor实例
                    System.Diagnostics.Debug.WriteLine("[DEBUG] 创建DocumentProcessor实例");
                    var documentProcessor = new DocumentProcessor((Word.Application)wordApplication);

                    // 获取action数组
                    if (!args.ContainsKey("action"))
                    {
                        return new ToolResult { Success = false, Error = "缺少action参数" };
                    }

                    System.Collections.IList actions = null;
                    var actionValue = args["action"];

                    if (actionValue is System.Collections.IList list)
                    {
                        actions = list;
                    }
                    else
                    {
                        return new ToolResult { Success = false, Error = "action参数必须是数组" };
                    }

                    var results = new List<object>();
                    var errors = new List<string>();

                    // // 检查操作是否与历史删除记录冲突
                    // var validationErrors = ValidateActionsAgainstHistory(actions);
                    // if (validationErrors.Count > 0)
                    // {
                    //     return new ToolResult
                    //     {
                    //         Success = false,
                    //         Error = $"操作验证失败: {string.Join("; ", validationErrors)}"
                    //     };
                    // }

                    // ===== 历史修改融合逻辑开始 =====
                    // 注意：每个action都独立处理，不在这里预先获取历史记录
                    // ===== 历史修改融合逻辑结束 =====

                    // 遍历每个操作并执行，每个action都按照顺序走一次9步
                    for (int i = 0; i < actions.Count; i++)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] ===== 第{i + 1}次操作开始 =====");

                        try
                        {
                            var action = actions[i] as Dictionary<string, object>;
                            if (action == null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作 {i + 1}: action格式错误，跳过");
                                errors.Add($"操作 {i + 1}: 格式错误");
                                continue;
                            }

                            // 提前检查操作类型
                            string operationType = null;
                            if (action.ContainsKey("type"))
                            {
                                operationType = action["type"]?.ToString();
                            }
                            else if (action.ContainsKey("action"))
                            {
                                operationType = action["action"]?.ToString();
                            }

                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作 {i + 1}: 类型={operationType}");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作 {i + 1} 预处理异常: {ex.Message}");
                            errors.Add($"操作 {i + 1}: 预处理异常 - {ex.Message}");
                            continue;
                        }



                        // ===== 每个action独立执行9步逻辑 =====

                        // 第1步：在每次循环开始时重新获取最新的历史记录

                        // 第1-1步：保存历史记录（在移除按钮之前）
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第1-1步：保存历史记录");

                        // 获取旧的历史记录（从DocumentState）
                        var (oldHistoricalDeletes, oldHistoricalInserts) = DocumentState.GetConsolidatedHistoricalRanges();

                        // 获取新的历史记录（从现有操作块中提取，去除按钮偏移的影响）
                        var (historicalDeletes, historicalInserts) = ExtractHistoryFromExistingBlocks(documentProcessor);

                        // 打印对比信息
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第1步：历史记录对比");
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 旧的历史记录（DocumentState）- 删除:{oldHistoricalDeletes.Count}个, 插入:{oldHistoricalInserts.Count}个");
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 新的历史记录（现有操作块）- 删除:{historicalDeletes.Count}个, 插入:{historicalInserts.Count}个");

                        // 打印详细的删除范围对比
                        if (oldHistoricalDeletes.Count > 0 || historicalDeletes.Count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 删除范围详细对比:");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 旧的删除范围: {string.Join(", ", oldHistoricalDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 新的删除范围: {string.Join(", ", historicalDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                        }

                        // 打印详细的插入范围对比
                        if (oldHistoricalInserts.Count > 0 || historicalInserts.Count > 0
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入范围详细对比:");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 旧的插入范围: {string.Join(", ", oldHistoricalInserts.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 新的插入范围: {string.Join(", ", historicalInserts.Select(r => $"[{r.Start},{r.End}]"))}");
                        }

                        // 第1-2步：移除待处理的按钮
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第1-2步：检查是否有新增丢弃这些按钮，有则需要删除");
                        int removedButtons = documentProcessor.RemoveAllPendingButtons();
                        bool buttonsWereRemoved = removedButtons > 0;
                        if (removedButtons == 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 没有发现待处理的按钮");
                        }

                        // 初始化变量
                        var currentActionRanges = (inserts: new List<(int Start, int End)>(), deletes: new List<(int Start, int End)>());

                        try
                        {
                            var action = actions[i] as Dictionary<string, object>;
                            if (action == null)
                            {
                                errors.Add($"操作 {i + 1}: 格式错误");
                                continue;
                            }

                            // 处理参数名映射（兼容不同的参数名和格式）
                            var normalizedAction = new Dictionary<string, object>();

                            // 首先确定操作类型
                            string operationType = null;
                            if (action.ContainsKey("type"))
                            {
                                operationType = action["type"]?.ToString();
                            }
                            else if (action.ContainsKey("action"))
                            {
                                // 处理错误的字段名"action"，应该是"type"
                                operationType = action["action"]?.ToString();
                            }

                            if (string.IsNullOrEmpty(operationType))
                            {
                                errors.Add($"操作 {i + 1}: 缺少操作类型(type字段)");
                                continue;
                            }

                            normalizedAction["type"] = operationType;

                            // 创建detail对象
                            var normalizedDetail = new Dictionary<string, object>();

                            // 处理detail字段（标准格式）
                            if (action.ContainsKey("detail"))
                            {
                                var detail = action["detail"] as Dictionary<string, object>;
                                if (detail != null)
                                {
                                    foreach (var kvp in detail)
                                    {
                                        // 将old_text映射为original
                                        if (kvp.Key == "old_text")
                                        {
                                            normalizedDetail["original"] = kvp.Value;
                                        }
                                        // 将new_text保持为new（DocumentProcessor期望的是"new"）
                                        else if (kvp.Key == "new_text")
                                        {
                                            normalizedDetail["new"] = kvp.Value;
                                        }
                                        else
                                        {
                                            normalizedDetail[kvp.Key] = kvp.Value;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                // 处理扁平格式（所有参数直接在action对象下）
                                foreach (var kvp in action)
                                {
                                    if (kvp.Key == "type" || kvp.Key == "action")
                                        continue; // 跳过type字段

                                    // 将old_text映射为original
                                    if (kvp.Key == "old_text")
                                    {
                                        normalizedDetail["original"] = kvp.Value;
                                    }
                                    // 将new_text保持为new（DocumentProcessor期望的是"new"）
                                    else if (kvp.Key == "new_text")
                                    {
                                        normalizedDetail["new"] = kvp.Value;
                                    }
                                    else
                                    {
                                        normalizedDetail[kvp.Key] = kvp.Value;
                                    }
                                }
                            }

                            normalizedAction["detail"] = normalizedDetail;

                            // 将操作对象序列化为JSON字符串
                            var serializer = new JavaScriptSerializer();
                            string jsonAction = serializer.Serialize(normalizedAction);

                            // 清理JSON字符串，确保它是有效的JSON
                            jsonAction = McpToolsHelpers.CleanJsonString(jsonAction);

                            // 第2步：验证、插入和获取完整范围
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第2步：验证、插入和获取完整范围");

                            // 第2-1步：验证start和end文本的有效性，获取删除操作的基础范围
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第2-1步：验证start和end文本的有效性");
                            var (validationPassed, validatedRanges) = ValidateAndGetRangesBeforeExecution(actions[i], (Word.Application)wordApplication, historicalDeletes, i + 1);
                            if (!validationPassed)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改验证失败，跳过此操作");

                                // 如果之前移除了按钮，需要恢复按钮
                                if (buttonsWereRemoved)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 验证失败，恢复之前移除的按钮");
                                    try
                                    {
                                        ApplyFormattingAndButtons(documentProcessor, i + 1);
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 按钮恢复完成");
                                    }
                                    catch (Exception restoreEx)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 按钮恢复失败: {restoreEx.Message}");
                                    }
                                }

                                continue;
                            }

                            // 保存验证阶段的原始位置，避免插入操作后Range对象被自动更新
                            var originalRanges = new Dictionary<string, (int Start, int End)>();
                            var originalValues = new Dictionary<string, object>(); // 用于存储非Range类型的值

                            foreach (var kvp in validatedRanges)
                            {
                                // Word Range是左闭右开的，转换为我们的双闭合格式
                                int closedStart = kvp.Value.Start;
                                int closedEnd = kvp.Value.End - 1; // 左闭右开 → 双闭合
                                originalRanges[kvp.Key] = (closedStart, closedEnd);
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 保存{kvp.Key}原始范围: Word范围[{kvp.Value.Start},{kvp.Value.End}] → 双闭合[{closedStart}, {closedEnd}]");
                            }

                            // 获取action detail以检查是否有position或cursor参数
                            var currentAction = actions[i] as Dictionary<string, object>;
                            var actionDetail = currentAction?.ContainsKey("detail") == true ? currentAction["detail"] as Dictionary<string, object> : null;
                            if (actionDetail != null)
                            {
                                if (actionDetail.ContainsKey("position"))
                                {
                                    originalValues["position"] = actionDetail["position"];
                                }
                                if (actionDetail.ContainsKey("cursor"))
                                {
                                    originalValues["cursor"] = actionDetail["cursor"];
                                }
                            }

                            // 第2-2步：执行文档操作，获取操作产生的范围
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第2-2步：执行文档操作");

                            (int Start, int End)? insertedRangeResult = null;
                            bool insertSuccess = true;

                            string currentActionType = currentAction?.ContainsKey("type") == true ? currentAction["type"]?.ToString() : null;

                            if (currentActionType?.ToLower() == "replace")
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 开始处理replace操作");

                                try
                                {
                                    // 对于replace操作，直接在end位置后插入新文本
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 获取actionDetail: {actionDetail != null}");

                                    string newText = actionDetail?.ContainsKey("new") == true ? actionDetail["new"]?.ToString() : null;

                                if (!string.IsNullOrEmpty(newText) && originalRanges.ContainsKey("end"))
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 准备执行replace，检查originalRanges");
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - originalRanges.ContainsKey('end'): {originalRanges.ContainsKey("end")}");
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - originalRanges.ContainsKey('start'): {originalRanges.ContainsKey("start")}");

                                    var (endStart, endEnd) = originalRanges["end"];
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - end范围: [{endStart}, {endEnd}]");

                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 准备创建Word.Range，位置: {endEnd + 1}");

                                    // 在end的最后一个位置后插入（end是双闭合，所以插入位置是endEnd + 1）
                                    Word.Range insertRangeLocal = ((Word.Application)wordApplication).ActiveDocument.Range(endEnd + 1, endEnd + 1);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - Word.Range创建成功");

                                    insertRangeLocal.Text = newText;

                                    // 获取实际插入范围 - 通过Range对象的Start/End属性（Word API返回左闭右开区间，转换为双闭区间）
                                    int closedStart = insertRangeLocal.Start;
                                    int closedEnd = insertRangeLocal.End - 1; // 左闭右开 → 双闭合
                                    insertedRangeResult = (closedStart, closedEnd);

                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 直接在位置{endEnd + 1}插入新文本，用word的API获取的实际范围(左闭右开): [{insertRangeLocal.Start}, {insertRangeLocal.End}]");
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 转换为双闭区间: [{closedStart}, {closedEnd}]");

                                    // 根据获取的实际范围重新找到文本并打印
                                    try
                                    {
                                        Word.Range verificationRange = ((Word.Application)wordApplication).ActiveDocument.Range(insertRangeLocal.Start, insertRangeLocal.End);
                                        string retrievedText = verificationRange.Text;
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 重新根据范围[{insertRangeLocal.Start}, {insertRangeLocal.End}]获取的文本: '{retrievedText}'");
                                    }
                                    catch (Exception verifyEx)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 重新获取文本时发生异常: {verifyEx.Message}");
                                    }
                                }
                                else
                                {
                                    insertSuccess = false;
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 替换操作缺少必要参数");
                                }
                                }
                                catch (Exception replaceEx)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 替换操作执行时发生异常: {replaceEx.Message}");
                                    insertSuccess = false;

                                    // 记录详细的错误信息
                                    if (replaceEx is System.OperationCanceledException)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 检测到OperationCanceledException，这通常表示Word操作被取消，可能是由于表格操作冲突");
                                    }
                                }
                            }
                            else if (currentActionType?.ToLower() == "insert")
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 开始处理insert操作");

                                try
                                {
                                    // 对于insert操作，在pre文本后插入新文本
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 获取actionDetail: {actionDetail != null}");

                                    string insertText = actionDetail?.ContainsKey("insert") == true ? actionDetail["insert"]?.ToString() : null;
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - insertText: '{insertText}'");

                                    // 检查插入方式并执行相应的插入逻辑
                                    bool hasInsertText = !string.IsNullOrEmpty(insertText);
                                    bool hasTraditionalInsert = originalRanges.ContainsKey("pre") || originalRanges.ContainsKey("post");
                                    bool hasPositionInsert = originalValues.ContainsKey("position");
                                    bool hasCursorInsert = originalValues.ContainsKey("cursor");

                                    if (hasInsertText)
                                    {
                                        int insertPosition = -1;

                                        // 确定插入位置
                                        if (hasTraditionalInsert && originalRanges.ContainsKey("pre"))
                                        {
                                            // 传统方式：在pre文本后插入
                                            var (preStart, preEnd) = originalRanges["pre"];
                                            insertPosition = preEnd + 1;
                                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 传统插入方式，插入位置: {insertPosition} (preEnd + 1)");
                                        }
                                        else if (hasPositionInsert)
                                        {
                                            // 直接位置插入
                                            insertPosition = (int)originalValues["position"];
                                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 直接位置插入，插入位置: {insertPosition}");
                                        }
                                        else if (hasCursorInsert)
                                        {
                                            // 光标位置插入
                                            try
                                            {
                                                // 获取当前光标位置
                                                Word.Selection selection = ((Word.Application)wordApplication).Selection;
                                                insertPosition = selection.Start;
                                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 光标位置插入，当前光标位置: {insertPosition}");
                                            }
                                            catch (Exception cursorEx)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取光标位置失败: {cursorEx.Message}");
                                                insertSuccess = false;
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            insertSuccess = false;
                                            System.Diagnostics.Debug.WriteLine($"[DEBUG] insert操作缺少有效的插入方式参数");
                                            continue;
                                        }

                                        // 验证插入位置是否有效
                                        int docLength = ((Word.Application)wordApplication).ActiveDocument.Characters.Count;
                                        if (insertPosition < 0 || insertPosition > docLength)
                                        {
                                            insertSuccess = false;
                                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入位置无效: {insertPosition}, 文档长度: {docLength}");
                                            continue;
                                        }

                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 准备创建Word.Range，插入位置: {insertPosition}");

                                        // 创建插入范围并执行插入
                                        Word.Range insertRangeLocal = ((Word.Application)wordApplication).ActiveDocument.Range(insertPosition, insertPosition);
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - Word.Range创建成功");

                                        int docLengthBefore = ((Word.Application)wordApplication).ActiveDocument.Characters.Count;
                                        int insertStart = insertRangeLocal.Start;
                                        insertRangeLocal.Text = insertText;
                                        int insertEnd = insertRangeLocal.End;
                                        int docLengthAfter = ((Word.Application)wordApplication).ActiveDocument.Characters.Count;

                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 文本插入成功:");
                    

                                        // 获取实际插入范围 - 通过Range对象的Start/End属性（Word API返回左闭右开区间，转换为双闭区间）
                                        int closedStart = insertRangeLocal.Start;
                                        int closedEnd = insertRangeLocal.End - 1; // 左闭右开 → 双闭合
                                        insertedRangeResult = (closedStart, closedEnd);

                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 直接在位置{insertPosition}插入新文本，用word的API获取的实际范围(左闭右开): [{insertRangeLocal.Start}, {insertRangeLocal.End}]");
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 转换为双闭区间: [{closedStart}, {closedEnd}]");

                                        // 根据获取的实际范围重新找到文本并打印
                                        try
                                        {
                                            Word.Range verificationRange = ((Word.Application)wordApplication).ActiveDocument.Range(insertRangeLocal.Start, insertRangeLocal.End);
                                            string retrievedText = verificationRange.Text;
                                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 重新根据范围[{insertRangeLocal.Start}, {insertRangeLocal.End}]获取的文本: '{retrievedText}'");
                                        }
                                        catch (Exception verifyEx)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 重新获取文本时发生异常: {verifyEx.Message}");
                                        }
                                    }
                                    else
                                    {
                                        insertSuccess = false;
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] insert操作缺少必要参数（insertText为空）");
                                    }
                                }
                                catch (Exception insertEx)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] insert操作执行时发生异常: {insertEx.Message}");
                                    insertSuccess = false;

                                    // 记录详细的错误信息
                                    if (insertEx is System.OperationCanceledException)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 检测到OperationCanceledException，这通常表示Word操作被取消，可能是由于表格操作冲突");
                                    }
                                }
                            }
                            else
                            {
                                // 对于其他不支持的操作类型
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次操作 - 不支持的操作类型: {currentActionType}");
                                insertSuccess = false;
                            }
                            if (!insertSuccess)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改操作失败，尝试重新应用格式和按钮");

                                // 即使操作失败，也要重新应用格式和按钮，因为第1步已经删除了所有按钮
                                try
                                {
                                    ApplyFormattingAndButtons(documentProcessor, i + 1);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改格式和按钮重新应用完成");
                                }
                                catch (Exception formatEx)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改格式和按钮重新应用失败: {formatEx.Message}");
                                }

                                errors.Add($"操作 {i + 1}: 操作失败");
                                continue;
                            }
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改操作成功");

                            // 获取操作的范围信息（只有insert操作才有插入范围）
                            var insertRange = insertedRangeResult.HasValue ? insertedRangeResult.Value : (Start: 0, End: -1);
                            if (insertedRangeResult.HasValue)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作产生的范围: [{insertRange.Start}, {insertRange.End}]");
                            }

                            // 2-3-4-3：综合计算最终的删除和插入范围
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第2-3-4-3步：综合计算最终的删除和插入范围");

                            // 检查操作类型
                            if (currentActionType?.ToLower() == "insert" && insertedRangeResult.HasValue)
                            {
                                // 对于insert操作，直接使用插入范围
                                currentActionRanges = (new List<(int Start, int End)> { insertRange }, new List<(int Start, int End)>());
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] Insert操作直接设置范围: inserts=[{insertRange.Start},{insertRange.End}]");
                            }
                            else
                            {
                                // 对于其他操作，使用原来的计算方法
                                currentActionRanges = CalculateFinalActionRanges(actions[i], originalRanges, insertRange);
                            }

                            if (currentActionRanges.inserts.Any() || currentActionRanges.deletes.Any())
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第4步：打印start和end之间选中的range");
                                if (currentActionRanges.deletes.Any())
                                {
                                    PrintRangeList($"第{i + 1}次新修改删除的range", currentActionRanges.deletes);
                                }
                                if (currentActionRanges.inserts.Any())
                                {
                                    PrintRangeList($"第{i + 1}次新修改新增的range", currentActionRanges.inserts);
                                }

                                // 第5步：得到序号，整理成[[1,2],[3,4]]这种形式
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第5步：得到序号，整理成数组形式");
                                if (currentActionRanges.deletes.Any())
                                {
                                    PrintRangeList($"第{i + 1}次新修改删除整理后的序号", currentActionRanges.deletes);
                                }
                                if (currentActionRanges.inserts.Any())
                                {
                                    PrintRangeList($"第{i + 1}次新修改新增整理后的序号", currentActionRanges.inserts);
                                }

                            // 第6步：存为最新的修改
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第6步：存为最新的修改");
                        }

                        // 初始化合并后的变量
                        List<(int Start, int End)> mergedDeletes = new List<(int Start, int End)>();
                        List<(int Start, int End)> mergedInserts = new List<(int Start, int End)>();

                        // 第6步：和历史记录对轴
                        if (historicalDeletes.Count > 0 || historicalInserts.Count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第6步：和历史记录对轴");
                            var (adjustedDeletes, adjustedInserts) = RealignHistoricalRanges(currentActionRanges.inserts, historicalDeletes, historicalInserts);

                            // 第7步：打印历史记录和对轴后的记录
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第7步：打印历史记录和对轴后的记录");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 历史删除记录: {string.Join(",", historicalDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 历史插入记录: {string.Join(",", historicalInserts.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对轴后删除记录: {string.Join(",", adjustedDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对轴后插入记录: {string.Join(",", adjustedInserts.Select(r => $"[{r.Start},{r.End}]"))}");

                            // 第7.5步：对消处理
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第7.5步：对消处理");

                            // 获取对消位置列表
                            var eliminationPositions = GetEliminationPositions(currentActionRanges.deletes, historicalInserts);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 需要对消的文档位置: {string.Join(",", eliminationPositions)}");

                            // 在文档中实际删除对消位置的文字
                            if (eliminationPositions.Count > 0)
                            {
                                documentProcessor.ApplyDocumentElimination(eliminationPositions);
                            }

                            // 进行序号级别的对消调整
                            var (eliminatedHistoricalDeletes, eliminatedHistoricalInserts, eliminatedNewDeletes, eliminatedNewInserts) = PerformOffsetElimination(adjustedDeletes, adjustedInserts, currentActionRanges.deletes, currentActionRanges.inserts);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对消后历史删除记录: {string.Join(",", eliminatedHistoricalDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对消后历史插入记录: {string.Join(",", eliminatedHistoricalInserts.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对消后最新修改删除记录: {string.Join(",", eliminatedNewDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对消后最新修改插入记录: {string.Join(",", eliminatedNewInserts.Select(r => $"[{r.Start},{r.End}]"))}");

                            // 第8步：和历史记录合并
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第8步：和历史记录合并");

                            // 合并删除列表
                            mergedDeletes = MergeRanges(eliminatedHistoricalDeletes, eliminatedNewDeletes);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 合并后删除记录: {string.Join(",", mergedDeletes.Select(r => $"[{r.Start},{r.End}]"))}");

                            // 合并插入列表
                            mergedInserts = MergeRanges(eliminatedHistoricalInserts, eliminatedNewInserts);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 合并后插入记录: {string.Join(",", mergedInserts.Select(r => $"[{r.Start},{r.End}]"))}");

                            // 检查并合并相连的区间
                            mergedDeletes = MergeConnectedRanges(mergedDeletes);
                            mergedInserts = MergeConnectedRanges(mergedInserts);

                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 最终合并后删除记录: {string.Join(",", mergedDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 最终合并后插入记录: {string.Join(",", mergedInserts.Select(r => $"[{r.Start},{r.End}]"))}");

                            // 这里应该更新历史记录，但现在暂时不更新，因为我们还没有持久化存储
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 新的历史修改记录已生成（暂时未持久化存储）");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第6步：历史记录为空，无需对轴，直接复制");

                            // 直接将当前修改作为合并结果（无需对消处理）
                            mergedDeletes = currentActionRanges.deletes;
                            mergedInserts = currentActionRanges.inserts;

                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 合并后删除记录: {string.Join(",", mergedDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 合并后插入记录: {string.Join(",", mergedInserts.Select(r => $"[{r.Start},{r.End}]"))}");
                        }

                        // 第8步：更新历史记录（将新修改融入历史记录）
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改-第8步：更新历史记录");

                        // 清除旧的历史记录
                        DocumentState.ClearAll();

                        // 将合并后的记录保存为新的历史记录
                        foreach (var deleteRange in mergedDeletes)
                        {
                            DocumentState.AddHistoricalDeleteRange(deleteRange.Start, deleteRange.End);
                        }
                        foreach (var insertRangeItem in mergedInserts)
                        {
                            DocumentState.AddHistoricalInsertRange(insertRangeItem.Start, insertRangeItem.End);
                        }
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 已将 {mergedDeletes.Count} 个删除记录和 {mergedInserts.Count} 个插入记录更新到历史");

                        // 第9步：根据更新后的历史记录进行格式和按钮操作
                        ApplyFormattingAndButtons(documentProcessor, i + 1);

                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{i + 1}次新修改执行完成");

                            results.Add(new
                            {
                                index = i + 1,
                                success = true,
                                action = action
                            });
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作 {i + 1} 执行异常: {ex.Message}");

                            // 如果之前移除了按钮，需要恢复按钮
                            if (buttonsWereRemoved)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作异常，恢复之前移除的按钮");
                                try
                                {
                                    ApplyFormattingAndButtons(documentProcessor, i + 1);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 按钮恢复完成");
                                }
                                catch (Exception restoreEx)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 按钮恢复失败: {restoreEx.Message}");
                                }
                            }

                            errors.Add($"操作 {i + 1}: {ex.Message}");
                            results.Add(new
                            {
                                index = i + 1,
                                success = false,
                                error = ex.Message,
                                action = actions[i]
                            });
                        }
                    }

                    await Task.CompletedTask;

                    // 计算成功操作数量
                    int successfulCount = 0;
                    foreach (var r in results)
                    {
                        var resultObj = r as dynamic;
                        if (resultObj != null && resultObj.success == true)
                        {
                            successfulCount++;
                        }
                    }

                    var result = new ToolResult
                    {
                        Success = errors.Count == 0,
                        Data = new
                        {
                            total_actions = actions.Count,
                            successful_actions = successfulCount,
                            failed_actions = errors.Count,
                            results = results,
                            errors = errors.Count > 0 ? errors : null
                        }
                    };

                    return result;
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"文档操作失败: {ex.Message}" };
                }
            };

            availableTools.Add(tool);
        }

        // /// <summary>
        // /// 验证操作是否与历史删除记录冲突
        // /// </summary>
        // /// <param name="actions">要验证的操作列表</param>
        // /// <returns>验证错误列表，如果为空则表示验证通过</returns>
        // private static List<string> ValidateActionsAgainstHistory(System.Collections.IList actions)
        // {
        //     var errors = new List<string>();
        //     var historicalDeleteRanges = DocumentState.HistoricalDeleteRanges;

        //     for (int i = 0; i < actions.Count; i++)
        //     {
        //         try
        //         {
        //             var action = actions[i] as Dictionary<string, object>;
        //             if (action == null || !action.ContainsKey("type"))
        //             {
        //                 continue;
        //             }

        //             string operationType = action["type"]?.ToString();
        //             var detail = action.ContainsKey("detail") ? action["detail"] as Dictionary<string, object> : null;

        //             if (detail == null)
        //             {
        //                 continue;
        //             }

        //             string error = null;

        //             switch (operationType?.ToLower())
        //             {
        //                 case "replace":
        //                     error = ValidateReplaceAction(detail, historicalDeleteRanges, i + 1);
        //                     break;

        //                 case "delete":
        //                     error = ValidateDeleteAction(detail, historicalDeleteRanges, i + 1);
        //                     break;

        //                 case "insert":
        //                     error = ValidateInsertAction(detail, historicalDeleteRanges, i + 1);
        //                     break;
        //             }

        //             if (!string.IsNullOrEmpty(error))
        //             {
        //                 errors.Add(error);
        //             }
        //         }
        //         catch (Exception ex)
        //         {
        //             errors.Add($"操作 {i + 1}: 验证过程中发生异常 - {ex.Message}");
        //         }
        //     }

        //     return errors;
        // }

        // /// <summary>
        // /// 验证替换操作是否与历史删除记录冲突
        // /// </summary>
        // private static string ValidateReplaceAction(Dictionary<string, object> detail, IReadOnlyList<(int Start, int End)> historicalDeletes, int operationIndex)
        // {
        //     if (!detail.ContainsKey("original"))
        //     {
        //         return null;
        //     }

        //     var original = detail["original"];

        //     try
        //     {
        //         // 获取Word应用程序实例
        //         var wordApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application") as Word.Application;
        //         if (wordApp == null || wordApp.ActiveDocument == null)
        //         {
        //             return $"操作 {operationIndex}: 无法访问Word文档";
        //         }

        //         // 对于字符串类型的original，检查start和end位置是否在删除范围内
        //         if (original is string originalText)
        //         {
        //             if (string.IsNullOrEmpty(originalText))
        //             {
        //                 return $"操作 {operationIndex}: 原文内容不能为空";
        //             }

        //             var documentProcessor = new DocumentProcessor(wordApp);
        //             Word.Range originalRange = documentProcessor.FindText(wordApp.ActiveDocument.Content, originalText);
        //             if (originalRange != null)
        //             {
        //                 // 检查start和end位置是否在历史删除记录范围内
        //                 foreach (var (deleteStart, deleteEnd) in historicalDeletes)
        //                 {
        //                     // 检查start位置是否在删除范围内（包含结束位置）
        //                     if (PositionInRange(originalRange.Start, deleteStart, deleteEnd))
        //                     {
        //                         return $"操作 {operationIndex}: 替换操作的start位置 {originalRange.Start} 位于已删除内容范围内 (位置 {deleteStart}-{deleteEnd})，无法执行替换操作";
        //                     }
        //                     // 检查end位置是否在删除范围内（包含结束位置）
        //                     if (PositionInRange(originalRange.End, deleteStart, deleteEnd))
        //                     {
        //                         return $"操作 {operationIndex}: 替换操作的end位置 {originalRange.End} 位于已删除内容范围内 (位置 {deleteStart}-{deleteEnd})，无法执行替换操作";
        //                     }
        //                 }
        //             }
        //         }
        //         else if (original is Dictionary<string, object> originalDict)
        //         {
        //             // 对于start/end格式的original，检查参数有效性
        //             string startText = originalDict.ContainsKey("start") ? originalDict["start"]?.ToString() : null;
        //             string endText = originalDict.ContainsKey("end") ? originalDict["end"]?.ToString() : null;


        //             // 检查start参数
        //             if (string.IsNullOrEmpty(startText))
        //             {
        //                 return $"操作 {operationIndex}: start参数不能为空";
        //             }

        //             // end参数可以为空（表示到文档结尾）
        //             if (endText == null)
        //             {
        //                 endText = "";
        //             }

        //             // 验证start和end文本在文档中存在
        //             var documentProcessor = new DocumentProcessor(wordApp);

        //             // 查找所有start文本匹配
        //             List<Word.Range> startMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, startText);
        //             if (startMatches.Count == 0)
        //             {
        //                 return $"操作 {operationIndex}: 无法在文档中找到start文本 '{startText}'";
        //             }

        //             // 查找所有end文本匹配（如果有end文本）
        //             List<Word.Range> endMatches = new List<Word.Range>();
        //             if (!string.IsNullOrEmpty(endText))
        //             {
        //                 endMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, endText);
        //                 if (endMatches.Count == 0)
        //                 {
        //                     return $"操作 {operationIndex}: 无法在文档中找到end文本 '{endText}'";
        //                 }
        //             }

        //             // 检查有效的start匹配数量
        //             int validStartCount = documentProcessor.GetValidMatchCount(startMatches, historicalDeletes);
        //             if (validStartCount == 0)
        //             {
        //                 return $"操作 {operationIndex}: 无法在文档中找到start文本 '{startText}'";
        //             }
        //             else if (validStartCount > 1)
        //             {
        //                 return $"操作 {operationIndex}: 找到多个有效的start文本匹配（{validStartCount}个），请提供更特异的start文本来唯一确定位置";
        //             }

        //             // 选择有效的start匹配
        //             Word.Range startRange = documentProcessor.SelectValidMatch(startMatches, historicalDeletes, requireUnique: true);

        //             // 检查有效的end匹配数量（如果有end文本）
        //             Word.Range endRange = null;
        //             if (!string.IsNullOrEmpty(endText))
        //             {
        //                 int validEndCount = documentProcessor.GetValidMatchCount(endMatches, historicalDeletes);
        //                 if (validEndCount == 0)
        //                 {
        //                     return $"操作 {operationIndex}: 无法在文档中找到end文本 '{endText}'";
        //                 }
        //                 else if (validEndCount > 1)
        //                 {
        //                     return $"操作 {operationIndex}: 找到多个有效的end文本匹配（{validEndCount}个），请提供更特异的end文本来唯一确定位置";
        //                 }

        //                 endRange = documentProcessor.SelectValidMatch(endMatches, historicalDeletes, requireUnique: true);
        //             }
        //         }
        //         else
        //         {
        //             return $"操作 {operationIndex}: original参数格式无效";
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         return $"操作 {operationIndex}: 验证替换操作时发生异常 - {ex.Message}";
        //     }

        //     return null;
        // }

        // /// <summary>
        // /// 验证删除操作是否与历史删除记录冲突
        // /// </summary>
        // private static string ValidateDeleteAction(Dictionary<string, object> detail, IReadOnlyList<(int Start, int End)> historicalDeletes, int operationIndex)
        // {
        //     if (!detail.ContainsKey("deletion"))
        //     {
        //         return null;
        //     }

        //     var deletion = detail["deletion"];
        //     Word.Range deletionRange = null;

        //     try
        //     {
        //         // 获取Word应用程序实例
        //         var wordApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application") as Word.Application;
        //         if (wordApp == null || wordApp.ActiveDocument == null)
        //         {
        //             return $"操作 {operationIndex}: 无法访问Word文档";
        //         }

        //         // 解析删除范围
        //         var documentProcessor = new DocumentProcessor(wordApp);

        //         if (deletion is string deletionText)
        //         {
        //             // 查找所有匹配，检查有效匹配数量
        //             List<Word.Range> matches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, deletionText);
        //             int validCount = documentProcessor.GetValidMatchCount(matches, historicalDeletes);
        //             if (validCount == 0)
        //             {
        //                 return $"操作 {operationIndex}: 无法在文档中找到要删除的文本 '{deletionText}'（该文本可能已被删除或不存在）";
        //             }
        //             else if (validCount > 1)
        //             {
        //                 return $"操作 {operationIndex}: 找到多个有效的删除文本匹配（{validCount}个），请提供更特异的文本来唯一确定要删除的内容";
        //             }

        //             deletionRange = documentProcessor.SelectValidMatch(matches, historicalDeletes, requireUnique: true);
        //         }
        //         else if (deletion is Dictionary<string, object> deletionDict)
        //         {
        //             string startText = deletionDict.ContainsKey("start") ? deletionDict["start"]?.ToString() : null;
        //             string endText = deletionDict.ContainsKey("end") ? deletionDict["end"]?.ToString() : null;

        //             // 检查start文本的有效匹配数量
        //             List<Word.Range> startMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, startText);
        //             int validStartCount = documentProcessor.GetValidMatchCount(startMatches, historicalDeletes);
        //             if (validStartCount == 0)
        //             {
        //                 return $"操作 {operationIndex}: 无法在文档中找到要删除的start文本 '{startText}'";
        //             }
        //             else if (validStartCount > 1)
        //             {
        //                 return $"操作 {operationIndex}: 删除操作找到多个有效的start文本匹配（{validStartCount}个），请提供更特异的start文本来唯一确定位置";
        //             }

        //             // 检查end文本的有效匹配数量
        //             List<Word.Range> endMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, endText);
        //             int validEndCount = documentProcessor.GetValidMatchCount(endMatches, historicalDeletes);
        //             if (validEndCount == 0)
        //             {
        //                 return $"操作 {operationIndex}: 无法在文档中找到要删除的end文本 '{endText}'";
        //             }
        //             else if (validEndCount > 1)
        //             {
        //                 return $"操作 {operationIndex}: 删除操作找到多个有效的end文本匹配（{validEndCount}个），请提供更特异的end文本来唯一确定位置";
        //             }

        //             Word.Range startRange = documentProcessor.SelectValidMatch(startMatches, historicalDeletes, requireUnique: true);
        //             Word.Range endRange = documentProcessor.SelectValidMatch(endMatches, historicalDeletes, requireUnique: true);

        //             if (startRange != null && endRange != null)
        //             {
        //                 // 创建范围
        //                 int selectionStart = Math.Min(startRange.Start, endRange.Start);
        //                 int selectionEnd = Math.Max(startRange.End, endRange.End);
        //                 deletionRange = wordApp.ActiveDocument.Range(selectionStart, selectionEnd);

        //                 // 检查删除范围是否与历史删除记录重叠
        //                 foreach (var (start, end) in historicalDeletes)
        //                 {
        //                     if (RangesOverlap(deletionRange.Start, deletionRange.End, start, end))
        //                     {
        //                         return $"操作 {operationIndex}: 要删除的文本 (位置 {deletionRange.Start}-{deletionRange.End}) 与已删除内容 (位置 {start}-{end}) 重叠，无法执行删除操作";
        //                     }
        //                 }
        //             }
        //         }

        //         if (deletionRange == null)
        //         {
        //             return $"操作 {operationIndex}: 无法找到有效的要删除的文本（所有匹配都在已删除内容范围内）";
        //         }

        //         // 检查删除范围是否与历史删除记录重叠
        //         foreach (var (start, end) in historicalDeletes)
        //         {
        //             if (RangesOverlap(deletionRange.Start, deletionRange.End, start, end))
        //             {
        //                 return $"操作 {operationIndex}: 要删除的文本 (位置 {deletionRange.Start}-{deletionRange.End}) 与已删除内容 (位置 {start}-{end}) 重叠，无法执行删除操作";
        //             }
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         return $"操作 {operationIndex}: 验证删除操作时发生异常 - {ex.Message}";
        //     }

        //     return null;
        // }

        // /// <summary>
        // /// 验证插入操作是否与历史删除记录冲突
        // /// </summary>
        // private static string ValidateInsertAction(Dictionary<string, object> detail, IReadOnlyList<(int Start, int End)> historicalDeletes, int operationIndex)
        // {
        //     // 检查pre和post上下文是否与历史删除记录重叠
        //     string pre = detail.ContainsKey("pre") ? detail["pre"]?.ToString() : null;
        //     string post = detail.ContainsKey("post") ? detail["post"]?.ToString() : null;

        //     if (string.IsNullOrEmpty(pre) && string.IsNullOrEmpty(post))
        //     {
        //         return null; // 如果都没有上下文，跳过验证
        //     }

        //     try
        //     {
        //         // 获取Word应用程序实例
        //         var wordApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application") as Word.Application;
        //         if (wordApp == null || wordApp.ActiveDocument == null)
        //         {
        //             return $"操作 {operationIndex}: 无法访问Word文档";
        //         }

        //         var documentProcessor = new DocumentProcessor(wordApp);

        //         // 检查pre上下文
        //         if (!string.IsNullOrEmpty(pre))
        //         {
        //             List<Word.Range> preMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, pre);
        //             int validPreCount = documentProcessor.GetValidMatchCount(preMatches, historicalDeletes);
        //             if (validPreCount == 0)
        //             {
        //                 return $"操作 {operationIndex}: 无法在文档中找到插入位置的上文 '{pre}'";
        //             }
        //             else if (validPreCount > 1)
        //             {
        //                 return $"操作 {operationIndex}: 插入位置找到多个有效的上文匹配（{validPreCount}个），请提供更特异的pre文本来唯一确定插入位置";
        //             }

        //             Word.Range preRange = documentProcessor.SelectValidMatch(preMatches, historicalDeletes, requireUnique: true);
        //             if (preRange != null)
        //             {
        //                 foreach (var (start, end) in historicalDeletes)
        //                 {
        //                     if (RangesOverlap(preRange.Start, preRange.End, start, end))
        //                     {
        //                         return $"操作 {operationIndex}: 插入位置的上文 '{pre}' (位置 {preRange.Start}-{preRange.End}) 与已删除内容 (位置 {start}-{end}) 重叠，无法执行插入操作";
        //                     }
        //                 }
        //             }
        //         }

        //         // 检查post上下文
        //         if (!string.IsNullOrEmpty(post))
        //         {
        //             List<Word.Range> postMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, post);
        //             int validPostCount = documentProcessor.GetValidMatchCount(postMatches, historicalDeletes);
        //             if (validPostCount == 0)
        //             {
        //                 return $"操作 {operationIndex}: 无法在文档中找到插入位置的下文 '{post}'";
        //             }
        //             else if (validPostCount > 1)
        //             {
        //                 return $"操作 {operationIndex}: 插入位置找到多个有效的下文匹配（{validPostCount}个），请提供更特异的post文本来唯一确定插入位置";
        //             }

        //             Word.Range postRange = documentProcessor.SelectValidMatch(postMatches, historicalDeletes, requireUnique: true);
        //             if (postRange != null)
        //             {
        //                 foreach (var (start, end) in historicalDeletes)
        //                 {
        //                     if (RangesOverlap(postRange.Start, postRange.End, start, end))
        //                     {
        //                         return $"操作 {operationIndex}: 插入位置的下文 '{post}' (位置 {postRange.Start}-{postRange.End}) 与已删除内容 (位置 {start}-{end}) 重叠，无法执行插入操作";
        //                     }
        //                 }
        //             }
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         return $"操作 {operationIndex}: 验证插入操作时发生异常 - {ex.Message}";
        //     }

        //     return null;
        // }

        // /// <summary>
        // /// 检查两个范围是否重叠
        // /// </summary>
        // private static bool RangesOverlap(int start1, int end1, int start2, int end2)
        // {
        //     return start1 < end2 && start2 < end1;
        // }

        // /// <summary>
        // /// 检查位置是否在指定范围内（包含结束位置）
        // /// </summary>
        // private static bool PositionInRange(int position, int rangeStart, int rangeEnd)
        // {
        //     return position >= rangeStart && position <= rangeEnd;
        // }

        /// <summary>
        /// 重新对轴历史修改记录以适应新的修改插入
        /// </summary>
        /// <param name="newInserts">新修改的插入区间列表（从小到大排列）</param>
        /// <param name="consolidatedDeletes">历史删除区间列表</param>
        /// <param name="consolidatedInserts">历史插入区间列表</param>
        /// <returns>重新对轴后的历史删除和插入区间</returns>
        private static (List<(int Start, int End)> adjustedDeletes, List<(int Start, int End)> adjustedInserts) RealignHistoricalRanges(
            List<(int Start, int End)> newInserts,
            List<(int Start, int End)> consolidatedDeletes,
            List<(int Start, int End)> consolidatedInserts)
        {
            // 合并所有历史范围（删除和插入），按起始位置排序
            var allHistoricalRanges = new List<(int Start, int End, bool IsInsert)>();
            allHistoricalRanges.AddRange(consolidatedDeletes.Select(r => (r.Start, r.End, false)));
            allHistoricalRanges.AddRange(consolidatedInserts.Select(r => (r.Start, r.End, true)));
            allHistoricalRanges.Sort((a, b) => a.Start.CompareTo(b.Start));

            var adjustedDeletes = new List<(int Start, int End)>();
            var adjustedInserts = new List<(int Start, int End)>();

            // 计算累积偏移量，处理每个历史范围
            int cumulativeOffset = 0;

            foreach (var historicalRange in allHistoricalRanges)
            {
                int rangeStart = historicalRange.Start;
                int rangeEnd = historicalRange.End;

                // 应用累积偏移
                int adjustedStart = rangeStart + cumulativeOffset;
                int adjustedEnd = rangeEnd + cumulativeOffset;

                // 检查是否有新插入影响这个范围
                bool rangeWasSplit = false;

                foreach (var newInsert in newInserts)
                {
                    int insertStart = newInsert.Start;
                    int insertLength = newInsert.End - newInsert.Start + 1;

                    // 如果插入点在历史范围内部，需要拆分
                    if (insertStart > rangeStart && insertStart <= rangeEnd)
                    {
                        // 第一个子范围：从历史起点到插入点之前
                        int firstPartEnd = insertStart - 1;
                        if (rangeStart <= firstPartEnd)
                        {
                            int finalFirstStart = rangeStart + cumulativeOffset;
                            int finalFirstEnd = firstPartEnd + cumulativeOffset;

                            if (historicalRange.IsInsert)
                            {
                                adjustedInserts.Add((finalFirstStart, finalFirstEnd));
                            }
                            else
                            {
                                adjustedDeletes.Add((finalFirstStart, finalFirstEnd));
                            }
                        }

                        // 第二个子范围：从插入点到历史终点，后续所有范围都要移动
                        int secondPartStart = insertStart + cumulativeOffset + insertLength;
                        int secondPartEnd = rangeEnd + cumulativeOffset + insertLength;

                        if (historicalRange.IsInsert)
                        {
                            adjustedInserts.Add((secondPartStart, secondPartEnd));
                        }
                        else
                        {
                            adjustedDeletes.Add((secondPartStart, secondPartEnd));
                        }

                        rangeWasSplit = true;
                        break; // 一个范围只能被一个插入拆分
                    }
                    // 如果插入点在历史范围起点之前，后续所有范围往后移动
                    else if (insertStart <= rangeStart)
                    {
                        adjustedStart += insertLength;
                        adjustedEnd += insertLength;
                    }
                }

                // 如果范围没有被拆分，添加调整后的完整范围
                if (!rangeWasSplit)
                {
                    if (historicalRange.IsInsert)
                    {
                        adjustedInserts.Add((adjustedStart, adjustedEnd));
                    }
                    else
                    {
                        adjustedDeletes.Add((adjustedStart, adjustedEnd));
                    }
                }
            }

            return (adjustedDeletes, adjustedInserts);
        }

        /// <summary>
        /// 在执行操作前验证文本有效性并获取范围
        /// </summary>
        /// <param name="singleAction">单个操作</param>
        /// <param name="wordApp">Word应用程序实例</param>
        /// <param name="historicalDeletes">历史删除记录</param>
        /// <param name="operationIndex">操作索引</param>
        /// <returns>验证结果和验证后的范围</returns>
        private static (bool validationPassed, Dictionary<string, Word.Range> validatedRanges) ValidateAndGetRangesBeforeExecution(
            object singleAction,
            Word.Application wordApp,
            IReadOnlyList<(int Start, int End)> historicalDeletes,
            int operationIndex)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作{operationIndex} - 开始验证ValidateAndGetRangesBeforeExecution");

            var action = singleAction as Dictionary<string, object>;
            if (action == null || !action.ContainsKey("type"))
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作{operationIndex} - action格式错误或缺少type字段");
                return (false, null);
            }

            string operationType = action["type"]?.ToString();
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作{operationIndex} - 操作类型: {operationType}");

            var detail = action.ContainsKey("detail") ? action["detail"] as Dictionary<string, object> : null;

            if (detail == null)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作{operationIndex} - detail为空");
                return (false, null);
            }

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作{operationIndex} - 创建DocumentProcessor");
            var documentProcessor = new DocumentProcessor(wordApp);
            var validatedRanges = new Dictionary<string, Word.Range>();

            // 根据操作类型进行验证
            switch (operationType?.ToLower())
            {
                case "replace":
                    object original = detail.ContainsKey("original") ? detail["original"] : null;
                    if (original is Dictionary<string, object> originalDict)
                    {
                        string startText = originalDict.ContainsKey("start") ? originalDict["start"]?.ToString() : null;
                        string endText = originalDict.ContainsKey("end") ? originalDict["end"]?.ToString() : null;

                        // 验证start文本
                        if (!string.IsNullOrEmpty(startText))
                        {
                            List<Word.Range> startMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, startText);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] start文本 '{startText}' 找到 {startMatches.Count} 个匹配");
                            for (int j = 0; j < startMatches.Count; j++)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] start匹配 {j + 1}: 位置[{startMatches[j].Start},{startMatches[j].End}]");
                            }

                            int validStartCount = documentProcessor.GetValidMatchCount(startMatches, historicalDeletes);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] start文本有效匹配数量: {validStartCount}");

                            if (validStartCount != 1)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] start文本验证失败，有效匹配数为{validStartCount}");
                                return (false, null);
                            }

                            Word.Range validStartRange = documentProcessor.SelectValidMatch(startMatches, historicalDeletes, false);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 选择有效的start匹配: 位置[{validStartRange.Start},{validStartRange.End}]");
                            validatedRanges["start"] = validStartRange;
                        }

                        // 验证end文本
                        if (!string.IsNullOrEmpty(endText))
                        {
                            List<Word.Range> endMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, endText);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] end文本 '{endText}' 找到 {endMatches.Count} 个匹配");
                            for (int j = 0; j < endMatches.Count; j++)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] end匹配 {j + 1}: 位置[{endMatches[j].Start},{endMatches[j].End}]");
                            }

                            int validEndCount = documentProcessor.GetValidMatchCount(endMatches, historicalDeletes);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] end文本有效匹配数量: {validEndCount}");

                            if (validEndCount != 1)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] end文本验证失败，有效匹配数为{validEndCount}");
                                return (false, null);
                            }

                            Word.Range validEndRange = documentProcessor.SelectValidMatch(endMatches, historicalDeletes, false);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 选择有效的end匹配: 位置[{validEndRange.Start},{validEndRange.End}]");
                            validatedRanges["end"] = validEndRange;
                        }
                    }
                    break;

                case "delete":
                    object deletion = detail.ContainsKey("deletion") ? detail["deletion"] : null;
                    if (deletion is Dictionary<string, object> deletionDict)
                    {
                        string startText = deletionDict.ContainsKey("start") ? deletionDict["start"]?.ToString() : null;
                        string endText = deletionDict.ContainsKey("end") ? deletionDict["end"]?.ToString() : null;

                        // 验证start文本
                        if (!string.IsNullOrEmpty(startText))
                        {
                            List<Word.Range> startMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, startText);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] start文本 '{startText}' 找到 {startMatches.Count} 个匹配");
                            for (int j = 0; j < startMatches.Count; j++)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] start匹配 {j + 1}: 位置[{startMatches[j].Start},{startMatches[j].End}]");
                            }

                            int validStartCount = documentProcessor.GetValidMatchCount(startMatches, historicalDeletes);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] start文本有效匹配数量: {validStartCount}");

                            if (validStartCount != 1)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] start文本验证失败，有效匹配数为{validStartCount}");
                                return (false, null);
                            }

                            Word.Range validStartRange = documentProcessor.SelectValidMatch(startMatches, historicalDeletes, false);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 选择有效的start匹配: 位置[{validStartRange.Start},{validStartRange.End}]");
                            validatedRanges["start"] = validStartRange;
                        }

                        // 验证end文本
                        if (!string.IsNullOrEmpty(endText))
                        {
                            List<Word.Range> endMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, endText);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] end文本 '{endText}' 找到 {endMatches.Count} 个匹配");
                            for (int j = 0; j < endMatches.Count; j++)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] end匹配 {j + 1}: 位置[{endMatches[j].Start},{endMatches[j].End}]");
                            }

                            int validEndCount = documentProcessor.GetValidMatchCount(endMatches, historicalDeletes);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] end文本有效匹配数量: {validEndCount}");

                            if (validEndCount != 1)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] end文本验证失败，有效匹配数为{validEndCount}");
                                return (false, null);
                            }

                            Word.Range validEndRange = documentProcessor.SelectValidMatch(endMatches, historicalDeletes, false);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 选择有效的end匹配: 位置[{validEndRange.Start},{validEndRange.End}]");
                            validatedRanges["end"] = validEndRange;
                        }
                    }
                    break;

                case "insert":
                    // 检查插入方式：传统方式(pre/post)、直接位置(position)、光标位置(cursor)
                    bool hasTraditionalInsert = detail.ContainsKey("pre") || detail.ContainsKey("post");
                    bool hasPositionInsert = detail.ContainsKey("position");
                    bool hasCursorInsert = detail.ContainsKey("cursor") && (bool)(detail["cursor"] ?? false);

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入方式检测 - 传统方式: {hasTraditionalInsert}, 位置插入: {hasPositionInsert}, 光标插入: {hasCursorInsert}");

                    // 确保只使用一种插入方式
                    int insertMethodCount = (hasTraditionalInsert ? 1 : 0) + (hasPositionInsert ? 1 : 0) + (hasCursorInsert ? 1 : 0);
                    if (insertMethodCount == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入操作缺少插入方式参数");
                        return (false, null);
                    }
                    if (insertMethodCount > 1)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入操作只能使用一种插入方式");
                        return (false, null);
                    }

                    // 传统方式：验证pre和post文本
                    if (hasTraditionalInsert)
                    {
                        string pre = detail.ContainsKey("pre") ? detail["pre"]?.ToString() : null;
                        string post = detail.ContainsKey("post") ? detail["post"]?.ToString() : null;

                        // 验证pre文本
                        if (!string.IsNullOrEmpty(pre))
                        {
                            List<Word.Range> preMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, pre);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] pre文本 '{pre}' 找到 {preMatches.Count} 个匹配");
                            for (int j = 0; j < preMatches.Count; j++)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] pre匹配 {j + 1}: 位置[{preMatches[j].Start},{preMatches[j].End}]");
                            }

                            int validPreCount = documentProcessor.GetValidMatchCount(preMatches, historicalDeletes);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] pre文本有效匹配数量: {validPreCount}");

                            if (validPreCount != 1)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] pre文本验证失败，有效匹配数为{validPreCount}");
                                return (false, null);
                            }

                            Word.Range validPreRange = documentProcessor.SelectValidMatch(preMatches, historicalDeletes, false);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 选择有效的pre匹配: 位置[{validPreRange.Start},{validPreRange.End}]");
                            validatedRanges["pre"] = validPreRange;
                        }

                        // 验证post文本
                        if (!string.IsNullOrEmpty(post))
                        {
                            List<Word.Range> postMatches = documentProcessor.FindAllText(wordApp.ActiveDocument.Content, post);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] post文本 '{post}' 找到 {postMatches.Count} 个匹配");
                            for (int j = 0; j < postMatches.Count; j++)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] post匹配 {j + 1}: 位置[{postMatches[j].Start},{postMatches[j].End}]");
                            }

                            int validPostCount = documentProcessor.GetValidMatchCount(postMatches, historicalDeletes);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] post文本有效匹配数量: {validPostCount}");

                            if (validPostCount != 1)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] post文本验证失败，有效匹配数为{validPostCount}");
                                return (false, null);
                            }

                            Word.Range validPostRange = documentProcessor.SelectValidMatch(postMatches, historicalDeletes, false);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 选择有效的post匹配: 位置[{validPostRange.Start},{validPostRange.End}]");
                            validatedRanges["post"] = validPostRange;
                        }
                    }
                    // 直接位置插入：验证position参数
                    else if (hasPositionInsert)
                    {
                        object positionObj = detail["position"];
                        if (positionObj is int position)
                        {
                            // 验证位置是否在文档范围内
                            int docLength = wordApp.ActiveDocument.Characters.Count;
                            if (position < 0 || position > docLength)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] position参数无效: {position}, 文档长度: {docLength}");
                                return (false, null);
                            }
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] position参数验证通过: {position}");

                            // position参数验证通过，不需要存储在validatedRanges中
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] position参数类型错误");
                            return (false, null);
                        }
                    }
                    // 光标位置插入：不需要验证，直接通过
                    else if (hasCursorInsert)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 光标位置插入，无需验证");
                        // cursor参数不需要存储在validatedRanges中
                    }
                    break;
            }

            return (true, validatedRanges);
        }

        /// <summary>
        /// 综合计算最终的删除和插入范围
        /// </summary>
        /// <param name="singleAction">单个操作</param>
        /// <param name="originalRanges">验证阶段保存的原始范围位置</param>
        /// <param name="insertRange">插入操作的范围</param>
        /// <returns>插入和删除的区间列表</returns>
        private static (List<(int Start, int End)> inserts, List<(int Start, int End)> deletes) CalculateFinalActionRanges(
            object singleAction,
            Dictionary<string, (int Start, int End)> originalRanges,
            (int Start, int End) insertRange)
        {
            var inserts = new List<(int Start, int End)>();
            var deletes = new List<(int Start, int End)>();

            if (originalRanges == null || originalRanges.Count == 0)
            {
                return (inserts, deletes);
            }

            var action = singleAction as Dictionary<string, object>;
            if (action == null || !action.ContainsKey("type"))
            {
                return (inserts, deletes);
            }

            string operationType = action["type"]?.ToString();
            var detail = action.ContainsKey("detail") ? action["detail"] as Dictionary<string, object> : null;

            if (detail == null)
            {
                return (inserts, deletes);
            }

            switch (operationType?.ToLower())
            {
                case "insert":
                    // 直接使用执行插入操作时获得的范围
                    inserts.Add(insertRange);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入操作范围: [{insertRange.Start}, {insertRange.End}]");
                    break;

                case "replace":
                    string newText = detail.ContainsKey("new") ? detail["new"]?.ToString() : null;

                    if (!string.IsNullOrEmpty(newText))
                    {
                        // 使用验证阶段保存的start和end原始范围来确定替换位置
                        if (originalRanges.ContainsKey("start") && originalRanges.ContainsKey("end"))
                        {
                            var (startStart, startEnd) = originalRanges["start"];
                            var (endStart, endEnd) = originalRanges["end"];

                        // 直接使用验证过的范围位置来创建删除范围
                        int startPos = Math.Min(startStart, endStart);
                        int endPos = Math.Max(startEnd, endEnd);

                        deletes.Add((startPos, endPos));
                        // 使用执行操作时通过Word API获得的实际插入范围，而不是数学计算
                        inserts.Add(insertRange);

                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 替换操作删除范围: [{startPos}, {endPos}]");
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 替换操作插入范围(通过API获得): [{insertRange.Start}, {insertRange.End}]");
                        }
                    }
                    break;

                case "delete":
                    // 使用验证阶段保存的start和end原始范围来确定删除位置
                    if (originalRanges.ContainsKey("start") && originalRanges.ContainsKey("end"))
                    {
                        var (startStart, startEnd) = originalRanges["start"];
                        var (endStart, endEnd) = originalRanges["end"];

                        // 直接使用验证过的范围位置来创建删除范围
                        int startPos = Math.Min(startStart, endStart);
                        int endPos = Math.Max(startEnd, endEnd);

                        deletes.Add((startPos, endPos));

                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 删除操作范围: [{startPos}, {endPos}]");
                    }
                    break;
            }

            return (inserts, deletes);
        }

        /// <summary>
        /// 从单个操作中提取插入和删除的区间范围
        /// </summary>
        /// <param name="singleAction">单个操作</param>
        /// <param name="wordApp">Word应用程序实例</param>
        /// <returns>插入和删除的区间列表</returns>
        private static (List<(int Start, int End)> inserts, List<(int Start, int End)> deletes) ExtractSingleActionRanges(
            object singleAction,
            Word.Application wordApp,
            IReadOnlyList<(int Start, int End)> historicalDeletes)
        {
            var inserts = new List<(int Start, int End)>();
            var deletes = new List<(int Start, int End)>();

            if (wordApp?.ActiveDocument == null)
            {
                return (inserts, deletes);
            }

            var documentProcessor = new DocumentProcessor(wordApp);
            var actions = new System.Collections.ArrayList { singleAction };

            return ExtractActionRanges(actions, wordApp);
        }

        /// <summary>
        /// 从操作列表中提取插入和删除的区间范围（简化版，不再做验证）
        /// </summary>
        /// <param name="actions">操作列表</param>
        /// <param name="wordApp">Word应用程序实例</param>
        /// <returns>插入和删除的区间列表</returns>
        private static (List<(int Start, int End)> inserts, List<(int Start, int End)> deletes) ExtractActionRanges(
            System.Collections.IList actions,
            Word.Application wordApp)
        {
            var inserts = new List<(int Start, int End)>();
            var deletes = new List<(int Start, int End)>();

            if (wordApp?.ActiveDocument == null)
            {
                return (inserts, deletes);
            }

            var documentProcessor = new DocumentProcessor(wordApp);

            foreach (var actionObj in actions)
            {
                try
                {
                    var action = actionObj as Dictionary<string, object>;
                    if (action == null || !action.ContainsKey("type"))
                    {
                        continue;
                    }

                    string operationType = action["type"]?.ToString();
                    var detail = action.ContainsKey("detail") ? action["detail"] as Dictionary<string, object> : null;

                    if (detail == null)
                    {
                        continue;
                    }

                    Word.Range range = null;

                    switch (operationType?.ToLower())
                    {
                        case "insert":
                            string pre = detail.ContainsKey("pre") ? detail["pre"]?.ToString() : null;
                            string post = detail.ContainsKey("post") ? detail["post"]?.ToString() : null;
                            string insertText = detail.ContainsKey("insert") ? detail["insert"]?.ToString() : null;

                            if (!string.IsNullOrEmpty(pre) && !string.IsNullOrEmpty(post) && !string.IsNullOrEmpty(insertText))
                            {
                                // 找到插入位置的范围
                                range = documentProcessor.FindTextByStartEnd(wordApp.ActiveDocument.Content, pre, post);
                                if (range != null)
                                {
                                    inserts.Add((range.Start, range.End));
                                }
                            }
                            break;

                        case "replace":
                            object original = detail.ContainsKey("original") ? detail["original"] : null;
                            string newText = detail.ContainsKey("new") ? detail["new"]?.ToString() : null;

                            if (original != null && !string.IsNullOrEmpty(newText))
                            {
                                // 获取原始文本的范围
                                if (original is string originalText)
                                {
                                    range = documentProcessor.FindText(wordApp.ActiveDocument.Content, originalText);
                                }
                                else if (original is Dictionary<string, object> originalDict)
                                {
                                    string startText = originalDict.ContainsKey("start") ? originalDict["start"]?.ToString() : null;
                                    string endText = originalDict.ContainsKey("end") ? originalDict["end"]?.ToString() : null;
                                    range = documentProcessor.FindTextByStartEnd(wordApp.ActiveDocument.Content, startText, endText);
                                }

                                if (range != null)
                                {
                                    deletes.Add((range.Start, range.End));
                                    // 插入范围在删除范围之后
                                    inserts.Add((range.End + 1, range.End + newText.Length));
                                }
                            }
                            break;

                        case "delete":
                            object deletion = detail.ContainsKey("deletion") ? detail["deletion"] : null;

                            if (deletion != null)
                            {
                                // 获取要删除文本的范围
                                if (deletion is string deletionText)
                                {
                                    range = documentProcessor.FindText(wordApp.ActiveDocument.Content, deletionText);
                                }
                                else if (deletion is Dictionary<string, object> deletionDict)
                                {
                                    string startText = deletionDict.ContainsKey("start") ? deletionDict["start"]?.ToString() : null;
                                    string endText = deletionDict.ContainsKey("end") ? deletionDict["end"]?.ToString() : null;
                                    range = documentProcessor.FindTextByStartEnd(wordApp.ActiveDocument.Content, startText, endText);
                                }

                                if (range != null)
                                {
                                    deletes.Add((range.Start, range.End));
                                }
                            }
                            break;
                    }
                }
                catch (Exception)
                {
                    // 忽略提取范围时的异常，继续处理其他操作
                }
            }

            // 排序区间
            inserts.Sort((a, b) => a.Start.CompareTo(b.Start));
            deletes.Sort((a, b) => a.Start.CompareTo(b.Start));

            return (inserts, deletes);
        }

        /// <summary>
        /// 打印范围列表的调试信息
        /// </summary>
        /// <param name="listName">列表名称</param>
        /// <param name="ranges">范围列表</param>
        private static void PrintRangeList(string listName, List<(int Start, int End)> ranges)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] {listName}: [{string.Join(",", ranges.Select(r => $"[{r.Start},{r.End}]"))}]");
        }

        /// <summary>
        /// 执行对消处理：当新修改删除与历史插入重合时，进行对消
        /// </summary>
        /// <param name="historicalDeletes">历史删除范围（已对轴）</param>
        /// <param name="historicalInserts">历史插入范围（已对轴）</param>
        /// <param name="newDeletes">新删除范围</param>
        /// <param name="newInserts">新插入范围</param>
        /// <returns>对消后的历史删除、历史插入、新删除、新插入范围</returns>
        private static (List<(int Start, int End)> eliminatedHistoricalDeletes, List<(int Start, int End)> eliminatedHistoricalInserts, List<(int Start, int End)> eliminatedNewDeletes, List<(int Start, int End)> eliminatedNewInserts) PerformOffsetElimination(
            List<(int Start, int End)> historicalDeletes,
            List<(int Start, int End)> historicalInserts,
            List<(int Start, int End)> newDeletes,
            List<(int Start, int End)> newInserts)
        {
            // 步骤1：获得对消的区间（新删除与历史插入的重合区域）
            var eliminationPositions = new List<int>();

            foreach (var newDelete in newDeletes)
            {
                foreach (var histInsert in historicalInserts)
                {
                    // 找出重合区域（闭区间）
                    int overlapStart = Math.Max(newDelete.Start, histInsert.Start);
                    int overlapEnd = Math.Min(newDelete.End, histInsert.End);

                    if (overlapStart <= overlapEnd)
                    {
                        // 将重合区域内的所有位置都加入对消列表
                        for (int pos = overlapStart; pos <= overlapEnd; pos++)
                        {
                            if (!eliminationPositions.Contains(pos))
                            {
                                eliminationPositions.Add(pos);
                            }
                        }
                    }
                }
            }

            // 对消位置排序
            eliminationPositions.Sort();

            // 将对消位置转换为区间格式进行显示
            var eliminationRanges = ConvertPositionsToRanges(eliminationPositions);
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对消区间: {string.Join(",", eliminationRanges.Select(r => $"[{r.Start},{r.End}]"))}");

            // 步骤2：在历史和新增修改中的删除和新增各个区间都找一遍
            var allRanges = new List<(int Start, int End, string Type)>();

            // 添加所有范围（按从小到大排列）
            foreach (var range in historicalDeletes) allRanges.Add((range.Start, range.End, "historical_delete"));
            foreach (var range in historicalInserts) allRanges.Add((range.Start, range.End, "historical_insert"));
            foreach (var range in newDeletes) allRanges.Add((range.Start, range.End, "new_delete"));
            foreach (var range in newInserts) allRanges.Add((range.Start, range.End, "new_insert"));

            allRanges.Sort((a, b) => a.Start.CompareTo(b.Start));

            // 步骤3：对消处理 - 删除对消位置上的内容，前移后续位置
            var resultHistoricalDeletes = new List<(int Start, int End)>();
            var resultHistoricalInserts = new List<(int Start, int End)>();
            var resultNewDeletes = new List<(int Start, int End)>();
            var resultNewInserts = new List<(int Start, int End)>();

            foreach (var range in allRanges)
            {
                var adjustedRange = ApplyEliminationToRange(range.Start, range.End, eliminationPositions);

                if (adjustedRange.HasValue)
                {
                    var (newStart, newEnd) = adjustedRange.Value;

                    switch (range.Type)
                    {
                        case "historical_delete":
                            resultHistoricalDeletes.Add((newStart, newEnd));
                            break;
                        case "historical_insert":
                            resultHistoricalInserts.Add((newStart, newEnd));
                            break;
                        case "new_delete":
                            resultNewDeletes.Add((newStart, newEnd));
                            break;
                        case "new_insert":
                            resultNewInserts.Add((newStart, newEnd));
                            break;
                    }
                }
                // 如果adjustedRange为null，说明整个范围都被对消了，跳过
            }

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 对消处理完成");

            return (resultHistoricalDeletes, resultHistoricalInserts, resultNewDeletes, resultNewInserts);
        }

        /// <summary>
        /// 对单个范围应用对消处理
        /// </summary>
        /// <param name="start">范围起始位置</param>
        /// <param name="end">范围结束位置</param>
        /// <param name="eliminationPositions">需要对消的位置列表（已排序）</param>
        /// <returns>对消后的范围，如果整个范围都被对消则返回null</returns>
        private static (int Start, int End)? ApplyEliminationToRange(int start, int end, List<int> eliminationPositions)
        {
            var remainingPositions = new List<int>();

            // 检查范围内的每个位置是否需要对消
            for (int pos = start; pos <= end; pos++)
            {
                if (!eliminationPositions.Contains(pos))
                {
                    remainingPositions.Add(pos);
                }
            }

            // 如果没有剩余位置，整个范围都被对消了
            if (remainingPositions.Count == 0)
            {
                return null;
            }

            // 重新计算范围的起始和结束位置
            int newStart = remainingPositions.Min();
            int newEnd = remainingPositions.Max();

            // 应用位置前移
            int offset = 0;
            foreach (int elimPos in eliminationPositions)
            {
                if (elimPos < newStart)
                {
                    offset++;
                }
                else
                {
                    break; // 对消位置在当前范围之后，不影响
                }
            }

            return (newStart - offset, newEnd - offset);
        }

        /// <summary>
        /// 将范围按对消区域分割，返回剩余的部分
        /// </summary>
        private static List<(int Start, int End)> SplitRangeByEliminationZones((int Start, int End) range, List<(int Start, int End)> eliminationZones)
        {
            var remainingParts = new List<(int Start, int End)> { range };

            foreach (var elimZone in eliminationZones)
            {
                var newParts = new List<(int Start, int End)>();

                foreach (var part in remainingParts)
                {
                    // 检查是否有重合
                    int overlapStart = Math.Max(part.Start, elimZone.Start);
                    int overlapEnd = Math.Min(part.End, elimZone.End);

                    if (overlapStart <= overlapEnd)
                    {
                        // 有重合，进行分割
                        // 左侧剩余部分
                        if (part.Start < overlapStart)
                        {
                            newParts.Add((part.Start, overlapStart - 1));
                        }
                        // 右侧剩余部分
                        if (part.End > overlapEnd)
                        {
                            newParts.Add((overlapEnd + 1, part.End));
                        }
                        // 重合部分被消除，不添加到newParts
                    }
                    else
                    {
                        // 没有重合，保持原样
                        newParts.Add(part);
                    }
                }

                remainingParts = newParts;
            }

            return remainingParts;
        }

        /// <summary>
        /// 对范围列表应用对消区域的偏移
        /// </summary>
        private static List<(int Start, int End)> ApplyOffsetToRanges(List<(int Start, int End)> ranges, List<(int Start, int End)> eliminationZones)
        {
            var result = new List<(int Start, int End)>();

            foreach (var range in ranges)
            {
                int offset = 0;

                // 计算这个范围之前的所有对消区域的总长度
                foreach (var elimZone in eliminationZones)
                {
                    if (elimZone.Start < range.Start)
                    {
                        // 对消区域完全在当前范围之前
                        offset += (elimZone.End - elimZone.Start + 1);
                    }
                    else if (elimZone.Start <= range.End && elimZone.End >= range.Start)
                    {
                        // 对消区域与当前范围有重合，只计算重合前面的部分
                        int effectiveEnd = Math.Min(elimZone.End, range.Start - 1);
                        if (effectiveEnd >= elimZone.Start)
                        {
                            offset += (effectiveEnd - elimZone.Start + 1);
                        }
                    }
                }

                result.Add((range.Start - offset, range.End - offset));
            }

            return result;
        }

        /// <summary>
        /// 操作块定义
        /// </summary>
        private class OperationBlock
        {
            public string Type { get; set; } // "delete", "insert", "replace"
            public int DeleteStart { get; set; } = -1;
            public int DeleteEnd { get; set; } = -1;
            public int InsertStart { get; set; } = -1;
            public int InsertEnd { get; set; } = -1;
        }

        /// <summary>
        /// 从现有的操作块中提取历史记录
        /// </summary>
        private static (List<(int Start, int End)> deletes, List<(int Start, int End)> inserts) ExtractHistoryFromExistingBlocks(DocumentProcessor documentProcessor)
        {
            var deletes = new List<(int Start, int End)>();
            var inserts = new List<(int Start, int End)>();

            try
            {
                // 通过反射获取Word应用程序
                var app = typeof(DocumentProcessor).GetField("_application", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                          ?.GetValue(documentProcessor) as Word.Application;
                if (app == null || app.ActiveDocument == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 无法获取Word应用程序或活动文档");
                    return (deletes, inserts);
                }

                Word.Document document = app.ActiveDocument;

                // 收集所有按钮的位置信息，用于计算偏移
                var buttonOffsets = new List<(int position, int length)>();
                foreach (Word.Bookmark bookmark in document.Bookmarks)
                {
                    string bookmarkName = bookmark.Name;
                    if (DocumentProcessor.OperationInfos.ContainsKey(bookmarkName))
                    {
                        try
                        {
                            // 估算按钮长度（通常包含操作类型和按钮文本）
                            int buttonLength = bookmark.Range.Text.Length;
                            if (buttonLength == 0) buttonLength = 20; // 默认估算长度
                            buttonOffsets.Add((bookmark.Range.Start, buttonLength));
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 发现按钮 '{bookmarkName}' 位置:{bookmark.Range.Start}, 长度:{buttonLength}");
                        }
                        catch
                        {
                            // 如果无法获取按钮长度，使用默认值
                            buttonOffsets.Add((bookmark.Range.Start, 20));
                        }
                    }
                }

                // 对按钮偏移按位置排序
                buttonOffsets = buttonOffsets.OrderBy(o => o.position).ToList();
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 收集到 {buttonOffsets.Count} 个按钮偏移信息");

                // 查找所有书签（按钮），提取操作块信息
                foreach (Word.Bookmark bookmark in document.Bookmarks)
                {
                    string bookmarkName = bookmark.Name;
                    if (DocumentProcessor.OperationInfos.ContainsKey(bookmarkName))
                    {
                        var operationInfo = DocumentProcessor.OperationInfos[bookmarkName] as dynamic;
                        if (operationInfo == null) continue;

                        string operationType = operationInfo.OperationType;
                        Word.Range operationRange = operationInfo.OperationRange;
                        Word.Range originalRange = operationInfo.OriginalRange;
                        Word.Range newRange = operationInfo.NewRange;

                        // 从操作块中提取范围信息，Word.Range是左闭右开的，需要转换为双闭区间
                        // 并减去按钮偏移的影响
                        if (operationType == "delete" && originalRange != null)
                        {
                            // Word Range是左闭右开的，转换为双闭合格式
                            int wordStart = originalRange.Start;
                            int wordEnd = originalRange.End;
                            int closedStart = wordStart;
                            int closedEnd = wordEnd - 1;

                            // 应用逆按钮偏移（减去按钮影响）
                            var correctedRange = ApplyReverseButtonOffsets(closedStart, closedEnd, buttonOffsets);
                            if (correctedRange.HasValue)
                            {
                                deletes.Add((correctedRange.Value.Start, correctedRange.Value.End));
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 从操作块提取删除范围: Word范围[{wordStart},{wordEnd}] -> 双闭合[{closedStart},{closedEnd}] -> 校正[{correctedRange.Value.Start},{correctedRange.Value.End}]");
                            }
                        }
                        else if (operationType == "insert" && operationRange != null)
                        {
                            // Word Range是左闭右开的，转换为双闭合格式
                            int wordStart = operationRange.Start;
                            int wordEnd = operationRange.End;
                            int closedStart = wordStart;
                            int closedEnd = wordEnd - 1;

                            // 应用逆按钮偏移（减去按钮影响）
                            var correctedRange = ApplyReverseButtonOffsets(closedStart, closedEnd, buttonOffsets);
                            if (correctedRange.HasValue)
                            {
                                inserts.Add((correctedRange.Value.Start, correctedRange.Value.End));
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 从操作块提取插入范围: Word范围[{wordStart},{wordEnd}] -> 双闭合[{closedStart},{closedEnd}] -> 校正[{correctedRange.Value.Start},{correctedRange.Value.End}]");
                            }
                        }
                        else if (operationType == "replace")
                        {
                            if (originalRange != null)
                            {
                                // Word Range是左闭右开的，转换为双闭合格式
                                int wordStart = originalRange.Start;
                                int wordEnd = originalRange.End;
                                int closedStart = wordStart;
                                int closedEnd = wordEnd - 1;

                                // 应用逆按钮偏移（减去按钮影响）
                                var correctedRange = ApplyReverseButtonOffsets(closedStart, closedEnd, buttonOffsets);
                                if (correctedRange.HasValue)
                                {
                                    deletes.Add((correctedRange.Value.Start, correctedRange.Value.End));
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 从操作块提取替换删除范围: Word范围[{wordStart},{wordEnd}] -> 双闭合[{closedStart},{closedEnd}] -> 校正[{correctedRange.Value.Start},{correctedRange.Value.End}]");
                                }
                            }
                            if (newRange != null)
                            {
                                // Word Range是左闭右开的，转换为双闭合格式
                                int wordStart = newRange.Start;
                                int wordEnd = newRange.End;
                                int closedStart = wordStart;
                                int closedEnd = wordEnd - 1;

                                // 应用逆按钮偏移（减去按钮影响）
                                var correctedRange = ApplyReverseButtonOffsets(closedStart, closedEnd, buttonOffsets);
                                if (correctedRange.HasValue)
                                {
                                    inserts.Add((correctedRange.Value.Start, correctedRange.Value.End));
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 从操作块提取替换插入范围: Word范围[{wordStart},{wordEnd}] -> 双闭合[{closedStart},{closedEnd}] -> 校正[{correctedRange.Value.Start},{correctedRange.Value.End}]");
                                }
                            }
                        }
                    }
                }

                // 合并重叠的范围
                deletes = MergeOverlappingRanges(deletes);
                inserts = MergeOverlappingRanges(inserts);

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 从现有操作块中提取的历史记录 - 删除: {deletes.Count}个, 插入: {inserts.Count}个");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取现有操作块的历史记录失败: {ex.Message}");
            }

            return (deletes, inserts);
        }

        /// <summary>
        /// 根据合并的历史记录生成操作块
        /// </summary>
        private static List<OperationBlock> GenerateOperationBlocks(List<(int Start, int End)> mergedDeletes, List<(int Start, int End)> mergedInserts)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 开始生成操作块 - 删除记录: {string.Join(",", mergedDeletes.Select(r => $"[{r.Start},{r.End}]"))}, 插入记录: {string.Join(",", mergedInserts.Select(r => $"[{r.Start},{r.End}]"))}");

            var operationBlocks = new List<OperationBlock>();

            // 1. 查找替换操作块：删除区间终点与新增区间起点相连
            var usedDeletes = new HashSet<(int Start, int End)>();
            var usedInserts = new HashSet<(int Start, int End)>();

            foreach (var deleteRange in mergedDeletes)
            {
                foreach (var insertRange in mergedInserts)
                {
                    // 检查是否相连：删除终点 + 1 = 新增起点
                    if (deleteRange.End + 1 == insertRange.Start)
                    {
                        var block = new OperationBlock
                        {
                            Type = "replace",
                            DeleteStart = deleteRange.Start,
                            DeleteEnd = deleteRange.End,
                            InsertStart = insertRange.Start,
                            InsertEnd = insertRange.End
                        };
                        operationBlocks.Add(block);
                        usedDeletes.Add(deleteRange);
                        usedInserts.Add(insertRange);
                        break; // 一个删除区间只能匹配一个插入区间
                    }
                }
            }

            // 2. 剩余的删除区间生成删除操作块
            foreach (var deleteRange in mergedDeletes)
            {
                if (!usedDeletes.Contains(deleteRange))
                {
                    var block = new OperationBlock
                    {
                        Type = "delete",
                        DeleteStart = deleteRange.Start,
                        DeleteEnd = deleteRange.End
                    };
                    operationBlocks.Add(block);
                }
            }

            // 3. 剩余的插入区间生成插入操作块
            foreach (var insertRange in mergedInserts)
            {
                if (!usedInserts.Contains(insertRange))
                {
                    var block = new OperationBlock
                    {
                        Type = "insert",
                        InsertStart = insertRange.Start,
                        InsertEnd = insertRange.End
                    };
                    operationBlocks.Add(block);
                }
            }

            // 打印生成的全部操作块
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 生成的操作块数量: {operationBlocks.Count}");
            foreach (var block in operationBlocks)
            {
                string desc = "";
                switch (block.Type)
                {
                    case "delete":
                        desc = $"删除: [{block.DeleteStart},{block.DeleteEnd}]";
                        break;
                    case "insert":
                        desc = $"插入: [{block.InsertStart},{block.InsertEnd}]";
                        break;
                    case "replace":
                        desc = $"替换: 删除[{block.DeleteStart},{block.DeleteEnd}] -> 插入[{block.InsertStart},{block.InsertEnd}]";
                        break;
                }
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作块: {desc}");
            }

            return operationBlocks;
        }

        /// <summary>
        /// 根据操作块更改文章格式并添加按钮
        /// </summary>
        private static void ApplyOperationBlocksToDocument(List<OperationBlock> operationBlocks, DocumentProcessor documentProcessor)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 开始应用 {operationBlocks.Count} 个操作块到文档");

            // 初始化按钮偏移跟踪列表（位置, 长度）
            var buttonOffsets = new List<(int position, int length)>();

            // 第一步：应用所有格式（删除线、背景色），并收集需要添加按钮的位置信息
            var buttonOperations = new List<(string operationType, Word.Range range, Word.Range originalRange, Word.Range newRange)>();

            foreach (var block in operationBlocks)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用操作块: {block.Type}");

                Word.Range operationRange = null;
                Word.Range originalRange = null;
                Word.Range newRange = null;

                switch (block.Type)
                {
                    case "delete":
                        // 删除操作：对删除范围添加删除线
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 删除操作 - 范围: [{block.DeleteStart},{block.DeleteEnd}]");
                        operationRange = GetDocumentRange(documentProcessor, block.DeleteStart, block.DeleteEnd);
                        if (operationRange != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取到删除范围，文本: '{operationRange.Text}'");
                            ApplyDeleteFormatting(operationRange);
                            // 收集按钮信息，稍后添加
                            buttonOperations.Add(("delete", operationRange, operationRange, null));
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 删除操作格式应用完成");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 无法获取删除范围");
                        }
                        break;

                    case "insert":
                        // 插入操作：对插入范围添加黄色背景
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入操作 - 范围: [{block.InsertStart},{block.InsertEnd}]");
                        operationRange = GetDocumentRange(documentProcessor, block.InsertStart, block.InsertEnd);
                        if (operationRange != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取到插入范围，文本: '{operationRange.Text}'");
                            ApplyInsertFormatting(operationRange);
                            // 收集按钮信息，稍后添加
                            buttonOperations.Add(("insert", operationRange, null, operationRange));
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入操作格式应用完成");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 无法获取插入范围");
                        }
                        break;

                    case "replace":
                        // 替换操作：删除部分加删除线，新增部分加黄色背景
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 替换操作 - 删除范围: [{block.DeleteStart},{block.DeleteEnd}], 插入范围: [{block.InsertStart},{block.InsertEnd}]");

                        if (block.DeleteStart >= 0 && block.DeleteEnd >= 0)
                        {
                            originalRange = GetDocumentRange(documentProcessor, block.DeleteStart, block.DeleteEnd);
                            if (originalRange != null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取到替换删除范围，文本: '{originalRange.Text}'");
                                ApplyDeleteFormatting(originalRange);
                            }
                        }

                        if (block.InsertStart >= 0 && block.InsertEnd >= 0)
                        {
                            newRange = GetDocumentRange(documentProcessor, block.InsertStart, block.InsertEnd);
                            if (newRange != null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取到替换插入范围，文本: '{newRange.Text}'");
                                ApplyInsertFormatting(newRange);
                            }
                        }

                        // 收集按钮信息，稍后添加（按钮应该在插入文本的末尾）
                        if (newRange != null)
                        {
                            buttonOperations.Add(("replace", newRange, originalRange, newRange));
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 替换操作格式应用完成");
                        }
                        else if (originalRange != null)
                        {
                            // 如果只有删除部分（不应该发生），放在删除文本末尾
                            buttonOperations.Add(("replace", originalRange, originalRange, newRange));
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 替换操作格式应用完成");
                        }
                        break;
                }
            }

            // 第二步：统一添加所有按钮，并记录偏移信息
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 开始添加 {buttonOperations.Count} 个操作按钮");

            // 通过反射获取Word应用程序，用于计算偏移
            var app = typeof(DocumentProcessor).GetField("_application", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                      ?.GetValue(documentProcessor) as Word.Application;

            foreach (var (operationType, range, originalRange, newRange) in buttonOperations)
            {
                try
                {
                    // 记录添加按钮前的文档长度
                    int docLengthBefore = app?.ActiveDocument?.Characters?.Count ?? 0;
                    int buttonPosition = range?.Start ?? 0;

                    // 添加按钮
                    documentProcessor.InsertInlineButtons(range, operationType, originalRange, newRange);

                    // 计算按钮添加后的长度变化
                    int docLengthAfter = app?.ActiveDocument?.Characters?.Count ?? 0;
                    int buttonLength = docLengthAfter - docLengthBefore;

                    if (buttonLength > 0)
                    {
                        buttonOffsets.Add((buttonPosition, buttonLength));
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 记录按钮偏移: 位置{buttonPosition}, 长度{buttonLength}");
                    }

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {operationType}操作按钮添加完成");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {operationType}操作按钮添加失败: {ex.Message}");
                }
            }

            // 第三步：计算最终范围（考虑按钮偏移后的实际位置）
            CalculateFinalRanges(buttonOffsets, documentProcessor);

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 所有操作块、按钮和最终范围计算完成");
        }

        /// <summary>
        /// 计算最终范围（考虑按钮偏移后的实际位置）
        /// </summary>
        private static void CalculateFinalRanges(List<(int position, int length)> buttonOffsets, DocumentProcessor documentProcessor)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 开始计算最终范围，按钮偏移数量: {buttonOffsets.Count}");

            // 清除旧的最终范围
            DocumentState.ClearFinalRanges();

            // 获取历史范围
            var (historicalDeletes, historicalInserts) = DocumentState.GetConsolidatedHistoricalRanges();

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 历史删除范围: {string.Join(",", historicalDeletes.Select(r => $"[{r.Start},{r.End}]"))}");
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 历史插入范围: {string.Join(",", historicalInserts.Select(r => $"[{r.Start},{r.End}]"))}");

            // 对每个历史范围应用按钮偏移
            foreach (var historicalRange in historicalDeletes)
            {
                var finalRange = ApplyButtonOffsetsToRange(historicalRange.Start, historicalRange.End, buttonOffsets);
                if (finalRange.HasValue)
                {
                    DocumentState.AddFinalDeleteRange(finalRange.Value.Start, finalRange.Value.End);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 最终删除范围: [{historicalRange.Start},{historicalRange.End}] -> [{finalRange.Value.Start},{finalRange.Value.End}]");
                }
            }

            foreach (var historicalRange in historicalInserts)
            {
                var finalRange = ApplyButtonOffsetsToRange(historicalRange.Start, historicalRange.End, buttonOffsets);
                if (finalRange.HasValue)
                {
                    DocumentState.AddFinalInsertRange(finalRange.Value.Start, finalRange.Value.End);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 最终插入范围: [{historicalRange.Start},{historicalRange.End}] -> [{finalRange.Value.Start},{finalRange.Value.End}]");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 最终范围计算完成");
        }

        /// <summary>
        /// 对单个范围应用逆按钮偏移（从包含按钮的范围恢复原始范围）
        /// </summary>
        private static (int Start, int End)? ApplyReverseButtonOffsets(int finalStart, int finalEnd, List<(int position, int length)> buttonOffsets)
        {
            int offset = 0;

            // 按位置排序按钮偏移（从小到大）
            var sortedOffsets = buttonOffsets.OrderBy(o => o.position).ToList();

            // 计算这个范围之前的所有按钮偏移（逆操作：减去偏移）
            foreach (var (buttonPos, buttonLen) in sortedOffsets)
            {
                if (buttonPos < finalStart)
                {
                    offset += buttonLen;
                }
                else
                {
                    // 按钮在范围之后或范围内，不影响这个范围的起始位置
                    break;
                }
            }

            // 应用逆偏移（减去偏移）
            int originalStart = finalStart - offset;
            int originalEnd = finalEnd - offset;

            // 验证范围有效性
            if (originalStart >= 0 && originalEnd >= originalStart)
            {
                return (originalStart, originalEnd);
            }

            return null;
        }

        /// <summary>
        /// 对单个范围应用按钮偏移
        /// </summary>
        private static (int Start, int End)? ApplyButtonOffsetsToRange(int originalStart, int originalEnd, List<(int position, int length)> buttonOffsets)
        {
            int offset = 0;

            // 按位置排序按钮偏移（从小到大）
            var sortedOffsets = buttonOffsets.OrderBy(o => o.position).ToList();

            // 计算这个范围之前的所有按钮偏移
            foreach (var (buttonPos, buttonLen) in sortedOffsets)
            {
                if (buttonPos <= originalStart)
                {
                    offset += buttonLen;
                }
                else
                {
                    // 按钮在范围之后，不影响这个范围
                    break;
                }
            }

            // 应用偏移
            int finalStart = originalStart + offset;
            int finalEnd = originalEnd + offset;

            // 验证范围有效性
            if (finalStart >= 0 && finalEnd >= finalStart)
            {
                return (finalStart, finalEnd);
            }

            return null;
        }

        /// <summary>
        /// 获取文档中的范围
        /// </summary>
        private static Word.Range GetDocumentRange(DocumentProcessor documentProcessor, int start, int end)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 尝试获取文档范围: 序号[{start},{end}]");

                // 通过反射获取DocumentProcessor的_application字段
                var app = typeof(DocumentProcessor).GetField("_application", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                          ?.GetValue(documentProcessor) as Word.Application;

                if (app?.ActiveDocument != null)
                {
                    var document = app.ActiveDocument;

                    // 获取文档的总字符数
                    int docLength = document.Characters.Count;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 文档总字符数: {docLength}");

                    // 确保序号在有效范围内
                    int actualStart = Math.Min(start, docLength - 1);
                    int actualEnd = Math.Min(end, docLength - 1);

                    if (actualStart < 0 || actualEnd < actualStart)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 序号范围无效: start={actualStart}, end={actualEnd}");
                        return null;
                    }

                    // Word.Range(start, end) 是左闭右开的，end不包含在内
                    // 对于我们的闭区间[start,end]（包含end），需要使用 Range(start, end + 1)
                    var range = document.Range(actualStart, actualEnd + 1);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取到文档范围: 位置[{actualStart},{actualEnd + 1}], 文本长度: {range.Characters.Count}, 文本: '{range.Text}'");

                    return range;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 无法获取Word应用程序或活动文档");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取文档范围失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 异常详情: {ex.StackTrace}");
            }

            return null;
        }

        /// <summary>
        /// 应用删除格式（删除线）
        /// </summary>
        private static void ApplyDeleteFormatting(Word.Range range)
        {
            try
            {
                for (int i = 1; i <= range.Characters.Count; i++)
                {
                    var charRange = range.Characters[i];
                    if (!string.IsNullOrWhiteSpace(charRange.Text) && charRange.Text != "\r" && charRange.Text != "\n")
                    {
                        charRange.Font.StrikeThrough = 1;
                    }
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
                // 逐字符设置背景色，避免Word自动扩展到整个段落
                for (int i = 1; i <= range.Characters.Count; i++)
                {
                    var charRange = range.Characters[i];
                    if (!string.IsNullOrWhiteSpace(charRange.Text) && charRange.Text != "\r" && charRange.Text != "\n")
                    {
                        charRange.Shading.BackgroundPatternColor = ConfigManager.DocumentColorHelper.InsertBackground;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用插入格式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 合并两个范围列表
        /// </summary>
        private static List<(int Start, int End)> MergeRanges(List<(int Start, int End)> list1, List<(int Start, int End)> list2)
        {
            var result = new List<(int Start, int End)>();
            result.AddRange(list1);
            result.AddRange(list2);
            return result;
        }

        /// <summary>
        /// 合并相连的范围（例如[1,2]和[3,4]合并为[1,4]）
        /// </summary>
        private static List<(int Start, int End)> MergeConnectedRanges(List<(int Start, int End)> ranges)
        {
            if (ranges.Count <= 1)
            {
                return ranges;
            }

            // 按起始位置排序
            var sortedRanges = ranges.OrderBy(r => r.Start).ToList();
            var merged = new List<(int Start, int End)> { sortedRanges[0] };

            for (int i = 1; i < sortedRanges.Count; i++)
            {
                var current = sortedRanges[i];
                var last = merged[merged.Count - 1];

                // 检查是否相连（包含端点相连的情况）
                if (current.Start <= last.End + 1)
                {
                    // 合并范围
                    merged[merged.Count - 1] = (last.Start, Math.Max(last.End, current.End));
                }
                else
                {
                    // 不相连，添加新范围
                    merged.Add(current);
                }
            }

            return merged;
        }

        /// <summary>
        /// 获取需要对消的文档位置列表
        /// </summary>
        /// <param name="newDeletes">新删除范围</param>
        /// <param name="historicalInserts">历史插入范围</param>
        /// <returns>需要对消的位置列表（已排序）</returns>
        private static List<int> GetEliminationPositions(List<(int Start, int End)> newDeletes, List<(int Start, int End)> historicalInserts)
        {
            var eliminationPositions = new List<int>();

            foreach (var newDelete in newDeletes)
            {
                foreach (var histInsert in historicalInserts)
                {
                    // 找出重合区域（闭区间）
                    int overlapStart = Math.Max(newDelete.Start, histInsert.Start);
                    int overlapEnd = Math.Min(newDelete.End, histInsert.End);

                    if (overlapStart <= overlapEnd)
                    {
                        // 将重合区域内的所有位置都加入对消列表
                        for (int pos = overlapStart; pos <= overlapEnd; pos++)
                        {
                            if (!eliminationPositions.Contains(pos))
                            {
                                eliminationPositions.Add(pos);
                            }
                        }
                    }
                }
            }

            // 对消位置排序
            eliminationPositions.Sort();

            return eliminationPositions;
        }

        /// <summary>
        /// 合并重叠的范围
        /// </summary>
        private static List<(int Start, int End)> MergeOverlappingRanges(List<(int Start, int End)> ranges)
        {
            if (ranges.Count == 0) return ranges;

            ranges.Sort((a, b) => a.Start.CompareTo(b.Start));
            var merged = new List<(int Start, int End)>();
            var current = ranges[0];

            for (int i = 1; i < ranges.Count; i++)
            {
                if (ranges[i].Start <= current.End + 1)
                {
                    current = (current.Start, Math.Max(current.End, ranges[i].End));
                }
                else
                {
                    merged.Add(current);
                    current = ranges[i];
                }
            }
            merged.Add(current);

            return merged;
        }
    }
}

