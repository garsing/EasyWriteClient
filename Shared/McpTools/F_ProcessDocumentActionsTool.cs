using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 处理句子操作工具
    /// 支持通过文本定位对文档进行句子级别的替换、删除、插入等操作（经 DocumentHost 解析渠道，Word/WPS）
    /// 写字：更新映射表 → 校验 detail.format → 写字并套显式字符格式（本工具不修改审阅开关）
    /// </summary>
    public static class F_ProcessDocumentActionsTool
    {
        /// <summary>
        /// 注册处理句子操作工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_process_document_actions"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_process_document_actions] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return resolveError;
                    }

                    Word.Document document = docHandle.Document;

                    string logPath = EasyWriteLog.BeginSession("process_document_actions");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 本次测试日志: {logPath}");
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_process_document_actions] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作前文档状态: 段落总数={document.Paragraphs.Count}");
                    FormatInheritHelper.DbgLogDocumentFontSample(document, "process_actions 入口");

                    if (!args.ContainsKey("action"))
                    {
                        System.Diagnostics.Debug.WriteLine("[ERROR] 缺少必需参数：action");
                        return new ToolResult { Success = false, Error = "缺少必需参数：action" };
                    }

                    // 第一步：更新句子与名称映射表（必须在解析参数之前，确保验证时能找到句子）
                    System.Diagnostics.Debug.WriteLine("[第一步] 更新句子与名称映射表");
                    FormatInheritHelper.DbgLog("第一步映射刷新前");
                    FormatInheritHelper.DbgLogDocumentFontSample(document, "第一步映射前");
                    try
                    {
                        string mappingError = await UpdateSentenceNameMapping(
                            document,
                            "process_actions_step1",
                            docHandle.DisallowBackendApi);
                        if (!string.IsNullOrEmpty(mappingError))
                        {
                            System.Diagnostics.Debug.WriteLine($"[ERROR] 更新映射表失败: {mappingError}");
                            return new ToolResult { Success = false, Error = mappingError };
                        }
                        System.Diagnostics.Debug.WriteLine(
                            $"[第一步完成] 映射表总数: {DocumentState.NameToContentMap.Count}");
                        FormatInheritHelper.DbgLogDocumentFontSample(document, "第一步映射后");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ERROR] 更新映射表异常: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"[ERROR] 异常堆栈: {ex.StackTrace}");
                        return new ToolResult { Success = false, Error = $"更新映射表失败: {ex.Message}" };
                    }

                    System.Diagnostics.Debug.WriteLine("[DEBUG] 开始解析操作参数");
                    var parseResult = ParseActions(args["action"], document);
                    bool parseSuccess = parseResult.Item1;
                    string parseError = parseResult.Item2;
                    List<ParagraphAction> actions = parseResult.Item3;
                    
                    if (!parseSuccess)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ERROR] 解析操作失败: {parseError}");
                        return new ToolResult { Success = false, Error = parseError };
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 解析操作成功，共 {actions.Count} 个操作");

                    for (int i = 0; i < actions.Count; i++)
                    {
                        actions[i].ActionIndex = i;
                    }

                    var ambiguitySettings = ConfigManager.GetDocumentActionsAmbiguitySettings();
                    var ambiguityAnchors = BuildAmbiguityAnchors(actions);
                    TableScopeIndex tableScope = TableSentenceScopeHelper.BuildIndexFromProcessCache(
                        DocumentState.LastProcessCache);
                    var ambiguityResult = SentenceCodeAmbiguityHelper.CheckBatch(
                        ambiguityAnchors,
                        DocumentState.Snapshot,
                        ambiguitySettings,
                        tableScope);
                    if (ambiguityResult.HasBlockingError)
                    {
                        return SentenceCodeAmbiguityHelper.BuildPreStep3FailureResult(
                            document.Name ?? "未命名文档",
                            actions.Count,
                            ambiguityResult);
                    }

                    ApplyResolvedDisplayPositions(actions, ambiguityResult);

                    string continuityError = ValidateAllActionsContinuity(actions, tableScope);
                    if (!string.IsNullOrEmpty(continuityError))
                    {
                        return SentenceCodeAmbiguityHelper.BuildPreStep3FailureResult(
                            document.Name ?? "未命名文档",
                            actions.Count,
                            "continuity_error",
                            continuityError);
                    }

                    // 第三步：直接修改文档（本工具不修改审阅开关）
                    System.Diagnostics.Debug.WriteLine("[第三步] 直接修改文档");
                    FormatInheritHelper.DbgLogDocumentFontSample(document, "第三步 modify 前");
                    List<ActionResult> modifyResults = await ModifyDocumentDirectly(
                        document,
                        actions,
                        tableScope);
                    FormatInheritHelper.DbgLogDocumentFontSample(document, "第三步 modify 后");
                    System.Diagnostics.Debug.WriteLine("[第三步完成]");

                    // 第四步：刷新映射表，解析 new_codes
                    System.Diagnostics.Debug.WriteLine("[第四步] 修改后刷新句子与名称映射表");
                    try
                    {
                        string postMappingError = await UpdateSentenceNameMapping(
                            document,
                            "process_actions_step4",
                            docHandle.DisallowBackendApi);
                        if (!string.IsNullOrEmpty(postMappingError))
                        {
                            System.Diagnostics.Debug.WriteLine($"[格式融合] 修改后刷新映射表 warning: {postMappingError}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[格式融合] 修改后刷新映射表异常: {ex.Message}");
                    }

                    ApplySnapshotDiffToResults(modifyResults);

                    var orderedResults = modifyResults.OrderBy(r => r.action_index).ToList();
                    object resultsPayload;
                    string batchMessage = null;

                    if (ShouldUseBulkSnapshotResults(orderedResults))
                    {
                        resultsPayload = BuildBulkSnapshotResultsObject(orderedResults);
                        batchMessage =
                            "本批操作已完成。由于含多项修改无法准确定位每项修改的影响，此处仅给出本批整体受影响的原句子编码与读回后的新句子。";
                    }
                    else
                    {
                        resultsPayload = orderedResults.Select(r => r.ToResponseObject()).ToList();
                    }

                    bool bulkMode = ShouldUseBulkSnapshotResults(orderedResults);
                    var navigateInputs = orderedResults
                        .Select(r => new NavigateActionInput
                        {
                            ActionIndex = r.action_index,
                            Type = r.type,
                            Success = r.success,
                            AffectedCodes = r.affected_codes,
                            NewCodes = r.new_codes,
                            NewSentences = r.new_sentences
                        })
                        .ToList();
                    var navigateSettings = ConfigManager.GetDocumentNavigateSettings();
                    // Desktop 常不注入 Application；从文档取，避免 WPS/空注入时强转失败
                    Word.Application wordApplicationTyped = wordApplication as Word.Application;
                    if (wordApplicationTyped == null)
                    {
                        try
                        {
                            wordApplicationTyped = document.Application;
                        }
                        catch (Exception)
                        {
                            wordApplicationTyped = null;
                        }
                    }

                    Word.Document activeDocument = document;
                    var navigateOutcome = DocumentNavigateHelper.BuildAndApply(
                        wordApplicationTyped,
                        activeDocument,
                        navigateInputs,
                        bulkMode,
                        navigateSettings.AutoNavigateAfterModify);
                    List<Dictionary<string, object>> navigates = navigateOutcome.Navigates;
                    Dictionary<string, object> navigateSummary = navigateOutcome.Summary;

                    bool allOk = orderedResults.All(r => r.success);
                    var result = new ToolResult
                    {
                        Success = allOk,
                        Data = new
                        {
                            document_name = document.Name ?? "未命名文档",
                            total_actions = actions.Count,
                            executed_actions = modifyResults.Count(r => r.success),
                            failed_actions = modifyResults.Count(r => !r.success),
                            message = batchMessage,
                            results = resultsPayload,
                            navigates = navigates,
                            navigate = navigateSummary
                        }
                    };

                    return result;
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ERROR] 处理句子操作异常: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[ERROR] 异常堆栈: {ex.StackTrace}");
                    return new ToolResult { Success = false, Error = $"处理句子操作失败: {ex.Message}" };
                }
            };
        }

        /// <summary>
        /// 解析操作列表：replace/delete/insert；定位对象为 codes+in_table / table / chart / image 四选一。
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
                                object original = detail.ContainsKey("original") ? detail["original"] : null;
                                if (!(original is Dictionary<string, object> originalDict))
                                {
                                    return (false, "replace 缺少 original 定位对象", null);
                                }

                                var originalLoc = ParseLocatorObject(originalDict, sentenceOnly: true, fieldLabel: "original");
                                if (!originalLoc.Success)
                                {
                                    return (false, originalLoc.Error, null);
                                }

                                action.SentenceNames = originalLoc.Spec.Codes;
                                action.OriginalScopeFields = new CodeAnchorScopeFields { TableId = originalLoc.Spec.InTableId };
                                
                                if (detail.ContainsKey("new"))
                                {
                                    action.Content = RemoveSentenceIdentifiers(detail["new"]?.ToString());
                                }

                                action.FormatSpec = ActionFormatHelper.ParseFormatObject(detail, required: true);
                                break;

                            case "delete":
                                object deletion = detail.ContainsKey("deletion") ? detail["deletion"] : null;
                                if (!(deletion is Dictionary<string, object> deletionDict))
                                {
                                    return (false, "delete 缺少 deletion 定位对象", null);
                                }

                                var deletionLoc = ParseLocatorObject(deletionDict, sentenceOnly: true, fieldLabel: "deletion");
                                if (!deletionLoc.Success)
                                {
                                    return (false, deletionLoc.Error, null);
                                }

                                action.SentenceNames = deletionLoc.Spec.Codes;
                                action.DeletionScopeFields = new CodeAnchorScopeFields { TableId = deletionLoc.Spec.InTableId };
                                break;

                            case "insert":
                                action.Position = "after";
                                if (detail.ContainsKey("insert"))
                                {
                                    action.Content = RemoveSentenceIdentifiers(detail["insert"]?.ToString());
                                }

                                if (detail.ContainsKey("target") && detail["target"] != null)
                                {
                                    if (!(detail["target"] is Dictionary<string, object> targetDict))
                                    {
                                        return (false, "insert.target 须为对象", null);
                                    }

                                    var targetLoc = ParseLocatorObject(targetDict, sentenceOnly: false, fieldLabel: "target");
                                    if (!targetLoc.Success)
                                    {
                                        return (false, targetLoc.Error, null);
                                    }

                                    action.TargetLocator = targetLoc.Spec;
                                    System.Diagnostics.Debug.WriteLine(
                                        $"[ProcessActions][insert] 解析 target Kind={targetLoc.Spec.Kind}");
                                }

                                action.FormatSpec = ActionFormatHelper.ParseFormatObject(detail, required: true);
                                break;
                        }

                        bool isValid = false;
                        if (action.Type?.ToLower() == "replace")
                        {
                            isValid = !string.IsNullOrEmpty(action.SentenceNames) && !string.IsNullOrEmpty(action.Content);
                        }
                        else if (action.Type?.ToLower() == "delete")
                        {
                            isValid = !string.IsNullOrEmpty(action.SentenceNames);
                        }
                        else if (action.Type?.ToLower() == "insert")
                        {
                            isValid = !string.IsNullOrEmpty(action.Content);
                        }

                        if (isValid)
                        {
                            actions.Add(action);
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[DEBUG] 解析操作失败：缺少必要字段，操作类型={action.Type}, SentenceNames={action.SentenceNames}, Target={action.TargetLocator?.Kind}, Content={action.Content}");
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

        private enum LocatorKind
        {
            Sentence,
            Table,
            Chart,
            Image
        }

        private sealed class LocatorSpec
        {
            public LocatorKind Kind { get; set; }
            public string Codes { get; set; }
            public string InTableId { get; set; }
            public string TableId { get; set; }
            public string ChartId { get; set; }
            public string ImageId { get; set; }
        }

        private sealed class LocatorParseResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public LocatorSpec Spec { get; set; }
        }

        /// <summary>
        /// 解析定位对象：委托 <see cref="InsertTargetLocatorHelper"/>（与 F_insert_svg_image 共用规则）。
        /// </summary>
        private static LocatorParseResult ParseLocatorObject(
            Dictionary<string, object> dict,
            bool sentenceOnly,
            string fieldLabel)
        {
            var shared = InsertTargetLocatorHelper.ParseLocatorObject(dict, sentenceOnly, fieldLabel);
            if (!shared.Success)
            {
                return new LocatorParseResult { Success = false, Error = shared.Error };
            }

            var s = shared.Spec;
            LocatorKind kind;
            switch (s.Kind)
            {
                case InsertTargetLocatorHelper.LocatorKind.Table:
                    kind = LocatorKind.Table;
                    break;
                case InsertTargetLocatorHelper.LocatorKind.Chart:
                    kind = LocatorKind.Chart;
                    break;
                case InsertTargetLocatorHelper.LocatorKind.Image:
                    kind = LocatorKind.Image;
                    break;
                default:
                    kind = LocatorKind.Sentence;
                    break;
            }

            return new LocatorParseResult
            {
                Success = true,
                Spec = new LocatorSpec
                {
                    Kind = kind,
                    Codes = s.Codes,
                    InTableId = s.InTableId,
                    TableId = s.TableId,
                    ChartId = s.ChartId,
                    ImageId = s.ImageId
                }
            };
        }

        private static List<DocumentActionAmbiguityAnchor> BuildAmbiguityAnchors(List<ParagraphAction> actions)
        {
            var anchors = new List<DocumentActionAmbiguityAnchor>();
            if (actions == null)
            {
                return anchors;
            }

            foreach (ParagraphAction action in actions)
            {
                string type = action.Type?.ToLower();
                if (type == "replace" && !string.IsNullOrEmpty(action.SentenceNames))
                {
                    anchors.Add(ToAmbiguityAnchor(action.ActionIndex, "original", action.SentenceNames, action.OriginalScopeFields));
                }
                else if (type == "delete" && !string.IsNullOrEmpty(action.SentenceNames))
                {
                    anchors.Add(ToAmbiguityAnchor(action.ActionIndex, "deletion", action.SentenceNames, action.DeletionScopeFields));
                }
                else if (type == "insert"
                         && action.TargetLocator != null
                         && action.TargetLocator.Kind == LocatorKind.Sentence
                         && !string.IsNullOrEmpty(action.TargetLocator.Codes))
                {
                    anchors.Add(ToAmbiguityAnchor(
                        action.ActionIndex,
                        "target",
                        action.TargetLocator.Codes,
                        new CodeAnchorScopeFields { TableId = action.TargetLocator.InTableId },
                        action.Content));
                }
            }

            return anchors;
        }

        private static DocumentActionAmbiguityAnchor ToAmbiguityAnchor(
            int actionIndex,
            string field,
            string codes,
            CodeAnchorScopeFields scopeFields,
            string insertContentPlaceholder = null)
        {
            scopeFields = scopeFields ?? new CodeAnchorScopeFields();
            return new DocumentActionAmbiguityAnchor
            {
                ActionIndex = actionIndex,
                Field = field,
                Codes = codes,
                TableId = scopeFields.TableId,
                InsertContentPlaceholder = insertContentPlaceholder
            };
        }

        private static void ApplyResolvedDisplayPositions(List<ParagraphAction> actions, AmbiguityCheckResult ambiguityResult)
        {
            if (actions == null || ambiguityResult?.ResolvedDisplayStartPositionByField == null)
            {
                return;
            }

            foreach (ParagraphAction action in actions)
            {
                string originalKey = SentenceCodeAmbiguityHelper.MakeFieldKey(action.ActionIndex, "original");
                if (ambiguityResult.ResolvedDisplayStartPositionByField.TryGetValue(originalKey, out int originalStart))
                {
                    action.ResolvedSentenceDisplayStart = originalStart;
                }

                string deletionKey = SentenceCodeAmbiguityHelper.MakeFieldKey(action.ActionIndex, "deletion");
                if (ambiguityResult.ResolvedDisplayStartPositionByField.TryGetValue(deletionKey, out int deletionStart))
                {
                    action.ResolvedSentenceDisplayStart = deletionStart;
                }

                string targetKey = SentenceCodeAmbiguityHelper.MakeFieldKey(action.ActionIndex, "target");
                if (ambiguityResult.ResolvedDisplayStartPositionByField.TryGetValue(targetKey, out int targetStart))
                {
                    action.ResolvedTargetDisplayStart = targetStart;
                }
            }
        }

        private static string ValidateAllActionsContinuity(List<ParagraphAction> actions, TableScopeIndex tableScope)
        {
            if (actions == null)
            {
                return null;
            }

            foreach (ParagraphAction action in actions)
            {
                string type = action.Type?.ToLower();
                if (type == "replace" || type == "delete")
                {
                    string error = ValidateSentenceSequenceContinuity(
                        action.SentenceNames,
                        action.ResolvedSentenceDisplayStart,
                        action.Type,
                        type == "replace"
                            ? action.OriginalScopeFields?.TableId
                            : action.DeletionScopeFields?.TableId,
                        tableScope);
                    if (!string.IsNullOrEmpty(error))
                    {
                        return error;
                    }
                }
                else if (type == "insert"
                         && action.TargetLocator != null
                         && action.TargetLocator.Kind == LocatorKind.Sentence)
                {
                    string error = ValidateSentenceSequenceContinuity(
                        action.TargetLocator.Codes,
                        action.ResolvedTargetDisplayStart,
                        action.Type + " (target)",
                        action.TargetLocator.InTableId,
                        tableScope);
                    if (!string.IsNullOrEmpty(error))
                    {
                        return error;
                    }
                }
            }

            return null;
        }

        private static void MoveSelectionAfterInsert(Word.Document doc, Word.Range insertRange)
        {
            if (doc?.Application == null || insertRange == null)
            {
                return;
            }

            WordRangeFinder.MoveSelectionToRange(
                doc.Application,
                insertRange,
                collapseToEnd: true,
                "insert");
        }

        private static void MoveSelectionAfterReplace(Word.Document doc, Word.Range replacedRange)
        {
            if (doc?.Application == null || replacedRange == null)
            {
                return;
            }

            WordRangeFinder.MoveSelectionToRange(
                doc.Application,
                replacedRange,
                collapseToEnd: true,
                "replace");
        }

        private static void MoveSelectionAfterDelete(Word.Document doc, int cursorPos)
        {
            if (doc?.Application?.Selection == null)
            {
                return;
            }

            try
            {
                int maxPos = Math.Max(0, doc.Content.End - 1);
                int pos = Math.Max(0, Math.Min(cursorPos, maxPos));
                doc.Application.Selection.SetRange(pos, pos);
                Word.Range scrollRange = doc.Range(pos, pos);
                WordRangeFinder.ScrollRangeIntoView(doc.Application, scrollRange, "delete");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[删除操作] 移动光标失败: {ex.Message}");
            }
        }

        private static Word.Table ResolveTableById(Word.Document doc, string tableId)
        {
            return SentenceCodeLocator.ResolveTableById(doc, tableId);
        }

        /// <summary>
        /// 第一步：更新句子与名称映射表（只更新现有文档的映射表，不处理新内容）
        /// 调用 WordDocumentExtractor.ProcessDocument 获取文章句子和名字的对应
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <returns>错误信息，如果成功则返回null或空字符串</returns>
        private static async Task<string> UpdateSentenceNameMapping(
            Word.Document doc,
            string snapshotSource,
            bool disallowBackendApi = false)
        {
            try
            {
                // 调用 WordDocumentExtractor.ProcessDocument 获取文章句子和名字的对应
                System.Diagnostics.Debug.WriteLine("[更新映射表] 调用 WordDocumentExtractor.ProcessDocument 获取文章句子和名字的对应");
                var options = ProcessDocumentOptions.ForProcessActions(snapshotSource);
                options.DisallowBackendApi = disallowBackendApi;
                var processingResult = WordDocumentExtractor.ProcessDocument(doc, options);
                System.Diagnostics.Debug.WriteLine(
                    $"[更新映射表] 文档处理完成，共 {processingResult.ChunkInfo.Count} 个chunk；" +
                    $"S_={DocumentState.SentenceNameMapping.Count} P_={DocumentState.ParagraphNameMapping.Count} " +
                    $"S→P索引={DocumentState.SentenceToParagraph.Count}");

                await Task.CompletedTask;
                return null; // 成功
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[更新映射表] 更新映射表失败: {ex.Message}");
                return $"更新映射表失败: {ex.Message}";
            }
        }

        /// <summary>
        /// 第三步：直接修改文档
        /// 基于句子名称序列定位句子并执行操作
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="actions">操作列表</param>
        /// <returns>操作结果列表</returns>
        private static async Task<List<ActionResult>> ModifyDocumentDirectly(
            Word.Document doc,
            List<ParagraphAction> actions,
            TableScopeIndex tableScope)
        {
            var results = new List<ActionResult>();

            try
            {
                // 按操作顺序处理（从后往前处理，避免位置偏移）
                var sortedActions = actions
                    .Select((action, originalIndex) => new { Action = action, OriginalIndex = originalIndex })
                    .ToList();

                // 先处理delete和replace（从后往前），再处理insert（从后往前）
                var deleteAndReplaceActions = sortedActions
                    .Where(x => x.Action.Type?.ToLower() == "delete" || x.Action.Type?.ToLower() == "replace")
                    .OrderByDescending(x => GetSentencePositionInDocument(
                        doc,
                        x.Action.SentenceNames,
                        x.Action.ResolvedSentenceDisplayStart,
                        x.Action.Type?.ToLower() == "replace"
                            ? x.Action.OriginalScopeFields?.TableId
                            : x.Action.DeletionScopeFields?.TableId,
                        tableScope))
                    .ToList();

                var insertActions = sortedActions
                    .Where(x => x.Action.Type?.ToLower() == "insert")
                    .OrderByDescending(x => GetInsertSortPosition(doc, x.Action, tableScope))
                    .ToList();

                // 先执行delete和replace操作
                foreach (var item in deleteAndReplaceActions)
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    var action = item.Action;
                    var result = new ActionResult
                    {
                        action_index = item.OriginalIndex,
                        type = action.Type,
                        target_index = 0
                    };

                    try
                    {
                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                            {
                                if (!TryPrepareActionFormat(
                                    doc, action, tableScope, result, out var formatSnapshot, out var formatMode))
                                {
                                    break;
                                }

                                var rep = ExecuteReplaceBySentenceNames(
                                    doc,
                                    action.SentenceNames,
                                    action.Content,
                                    action.ResolvedSentenceDisplayStart,
                                    action.OriginalScopeFields?.TableId,
                                    tableScope,
                                    formatSnapshot,
                                    formatMode);
                                FillReplaceActionResult(result, rep);
                                break;
                            }

                            case "delete":
                            {
                                var del = ExecuteDeleteBySentenceNames(
                                    doc,
                                    action.SentenceNames,
                                    action.ResolvedSentenceDisplayStart,
                                    action.DeletionScopeFields?.TableId,
                                    tableScope);
                                FillDeleteActionResult(result, del);
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        result.success = false;
                        result.message = $"操作失败: {ex.Message}";
                    }

                    results.Add(result);
                }

                // 再执行insert操作
                foreach (var item in insertActions)
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    var action = item.Action;
                    var result = new ActionResult
                    {
                        action_index = item.OriginalIndex,
                        type = action.Type,
                        target_index = 0
                    };

                    try
                    {
                        if (!TryPrepareActionFormat(
                            doc, action, tableScope, result, out var formatSnapshot, out var formatMode))
                        {
                            results.Add(result);
                            continue;
                        }

                        var ins = ExecuteInsert(
                            doc,
                            action.TargetLocator,
                            action.Content,
                            action.ResolvedTargetDisplayStart,
                            tableScope,
                            formatSnapshot,
                            formatMode);
                        FillInsertActionResult(result, ins, action.TargetLocator);
                    }
                    catch (Exception ex)
                    {
                        result.success = false;
                        result.message = $"操作失败: {ex.Message}";
                    }

                    results.Add(result);
                }

                // 按原始顺序排序
                results = results.OrderBy(r => r.action_index).ToList();

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[直接修改文档] 修改文档失败: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// 根据句子名称序列获取在文档中的位置（用于排序）
        /// </summary>
        private static int GetSentencePositionInDocument(
            Word.Document doc,
            string sentenceNames,
            int? displayStartPosition = null,
            string tableId = null,
            TableScopeIndex tableScope = null)
        {
            if (string.IsNullOrEmpty(sentenceNames) || !displayStartPosition.HasValue)
            {
                return 0;
            }

            try
            {
                var names = ParseCodeList(sentenceNames);
                if (names.Count == 0 || !names[0].StartsWith("S_", StringComparison.Ordinal))
                {
                    return 0;
                }

                Word.Range span = DisplayPositionRangeResolver.ResolveSpan(
                    doc,
                    names,
                    displayStartPosition.Value,
                    tableId,
                    tableScope,
                    debugTag: "process_actions_sort");
                if (span != null)
                {
                    return span.Start;
                }
            }
            catch
            {
            }

            return 0;
        }

        /// <summary>
        /// 根据 S_ 编码序列执行替换操作（replace 不支持 T_）
        /// </summary>
        private static bool TryPrepareActionFormat(
            Word.Document doc,
            ParagraphAction action,
            TableScopeIndex tableScope,
            ActionResult result,
            out Dictionary<string, object> snapshot,
            out string formatMode)
        {
            snapshot = null;
            formatMode = null;

            ActionFormatHelper.Spec spec = action?.FormatSpec;
            if (spec == null || !spec.Ok)
            {
                result.success = false;
                result.message = spec?.Error ?? ActionFormatHelper.MissingFormatError;
                return false;
            }

            if (!ActionFormatHelper.TryBuildSnapshot(
                doc, spec, tableScope, out snapshot, out formatMode, out string error))
            {
                result.success = false;
                result.message = error ?? ActionFormatHelper.MissingFormatError;
                return false;
            }

            return true;
        }

        private static ReplaceExecutionResult ExecuteReplaceBySentenceNames(
            Word.Document doc,
            string sentenceNames,
            string newContent,
            int? displayStartPosition = null,
            string tableId = null,
            TableScopeIndex tableScope = null,
            Dictionary<string, object> formatSnapshot = null,
            string formatMode = null)
        {
            var result = new ReplaceExecutionResult { Success = false, NewSentences = new List<string>() };

            FormatInheritHelper.DbgLog($"ExecuteReplace START codes=\"{sentenceNames}\"");

            try
            {
                if (string.IsNullOrEmpty(sentenceNames))
                {
                    result.ErrorMessage = "缺少 original.codes";
                    return result;
                }

                if (!displayStartPosition.HasValue)
                {
                    result.ErrorMessage = "缺少唯一定位结果";
                    return result;
                }

                var names = ParseCodeList(sentenceNames);
                if (names.Count == 0)
                {
                    result.ErrorMessage = "original.codes 为空";
                    return result;
                }

                if (names.Any(n => n.StartsWith("T_", StringComparison.Ordinal)
                    || n.StartsWith("C_", StringComparison.Ordinal)))
                {
                    result.ErrorMessage = "replace 仅支持 S_ 编码；T_/C_ 仅用于 insert 锚点";
                    return result;
                }

                result.OriginalCodes = names;
                FormatInheritHelper.DbgLog($"ExecuteReplace S_ count={names.Count} codes=[{string.Join(",", names)}]");

                Word.Range targetRange = DisplayPositionRangeResolver.ResolveSpan(
                    doc,
                    names,
                    displayStartPosition.Value,
                    tableId,
                    tableScope,
                    debugTag: "process_actions_replace");

                if (targetRange == null)
                {
                    result.ErrorMessage = "未找到任何句子 Range";
                    System.Diagnostics.Debug.WriteLine($"[替换操作] 未找到任何句子Range");
                    return result;
                }

                string originalRangeText = targetRange.Text ?? "";
                FormatInheritHelper.DbgLogRangeFonts("ExecuteReplace targetRange BEFORE Text=", targetRange);

                // 大模型输入后、分句/编码前：新文段尾 \r 数量与原文对齐
                string newContentBeforeAlign = newContent;
                newContent = TextUtils.EnsureReplaceContentTrailingParagraphMarks(newContent, originalRangeText);
                if (newContent != newContentBeforeAlign)
                {
                    string oldSuffix = TextUtils.GetTrailingParagraphSuffix(originalRangeText);
                    int newSuffixBefore = TextUtils.CountTrailingParagraphMarks(newContentBeforeAlign);
                    FormatInheritHelper.DbgLog(
                        $"ExecuteReplace 段尾对齐: 原文 suffix len={oldSuffix.Length} " +
                        $"({TextUtils.EscapeNewlinesForDb(oldSuffix)}), " +
                        $"新文对齐前 len={newSuffixBefore} → 对齐后 len={oldSuffix.Length}");
                }

                List<string> newSentences = SentenceSplitter.SplitForActionContent(newContent);
                if (newSentences.Count == 0)
                {
                    result.ErrorMessage = "新内容无法切分为句子";
                    System.Diagnostics.Debug.WriteLine($"[替换操作] 新内容无法切分为句子");
                    return result;
                }

                result.NewSentences = newSentences;

                if (formatSnapshot == null || formatSnapshot.Count == 0 || string.IsNullOrEmpty(formatMode))
                {
                    result.ErrorMessage = ActionFormatHelper.MissingFormatError;
                    return result;
                }

                FormatInheritHelper.DbgLog(
                    $"ExecuteReplace newSentenceCount={newSentences.Count} codes={names.Count} " +
                    $"format_mode={formatMode}");

                Dictionary<string, object> pendingPreReplaceSnapshot = null;
                if (names.Count >= 3)
                {
                    pendingPreReplaceSnapshot = FormatInheritHelper.ExtractSnapshot(targetRange);
                    result.PendingPreReplaceSnapshot = pendingPreReplaceSnapshot;
                    FormatContextHelper.DbgLog(
                        $"ExecuteReplace Extract 暂存 pre_replace 首句={names[0]} " +
                        $"fingerprint={FormatInheritHelper.BuildFingerprint(pendingPreReplaceSnapshot)}");
                }

                var newContentBuilder = new StringBuilder();
                foreach (var sentence in newSentences)
                {
                    newContentBuilder.Append(sentence);
                }

                string trailingSuffix = TextUtils.GetTrailingParagraphSuffix(originalRangeText);
                int trailingLen = trailingSuffix.Length;

                string replacementText = newContentBuilder.ToString();
                if (trailingLen > 0)
                {
                    replacementText = TextUtils.TrimTrailingParagraphMarks(replacementText);
                    int rangeSpan = targetRange.End - targetRange.Start;
                    if (rangeSpan > trailingLen)
                    {
                        targetRange = doc.Range(targetRange.Start, targetRange.End - trailingLen);
                        FormatInheritHelper.DbgLog(
                            $"ExecuteReplace 保留文档段尾: len={trailingLen} " +
                            $"({TextUtils.EscapeNewlinesForDb(trailingSuffix)})，仅替换正文");
                    }
                }

                FormatInheritHelper.DbgLog(
                    WordNoteMarkHelper.ContainsNoteMarks(replacementText)
                        ? "ExecuteReplace 即将写入（含 ^f/^e → 脚注/尾注引用）"
                        : "ExecuteReplace 即将 targetRange.Text=...（会清除该区间直接格式）");
                Word.Range writtenRange = WordNoteMarkHelper.WriteTextWithNoteMarks(
                    doc, targetRange, replacementText);
                targetRange = writtenRange ?? targetRange;
                FormatInheritHelper.DbgLogRangeFonts("ExecuteReplace targetRange AFTER Text=", targetRange);

                List<string> sentencesForFormat = trailingLen > 0
                    ? newSentences.ConvertAll(TextUtils.TrimTrailingParagraphMarks)
                    : newSentences;
                Word.Range newRange = doc.Range(targetRange.Start, targetRange.End);

                try
                {
                    FormatInheritHelper.DbgLog($"ExecuteReplace ApplySnapshot format_mode={formatMode}");
                    FormatInheritHelper.ApplySnapshotToSentences(
                        doc, newRange, sentencesForFormat, formatSnapshot, FindSentenceInRange);
                    result.FormatMode = formatMode;
                    FormatInheritHelper.DbgLogRangeFonts("ExecuteReplace targetRange AFTER format", newRange);
                }
                catch (Exception ex)
                {
                    result.FormatApplyError = ex.Message;
                    result.ErrorMessage = $"文字已写入，但字符格式套用失败: {ex.Message}";
                    System.Diagnostics.Debug.WriteLine($"[格式] format_apply_error: {ex.Message}");
                    return result;
                }

                result.Success = true;
                MoveSelectionAfterReplace(doc, targetRange);
                System.Diagnostics.Debug.WriteLine($"[替换操作] 已替换 {names.Count} 个句子");
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                System.Diagnostics.Debug.WriteLine($"[替换操作] 替换失败: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// 根据 S_ 编码序列执行删除操作（delete 不支持 T_）
        /// </summary>
        private static DeleteExecutionResult ExecuteDeleteBySentenceNames(
            Word.Document doc,
            string sentenceNames,
            int? displayStartPosition = null,
            string tableId = null,
            TableScopeIndex tableScope = null)
        {
            var result = new DeleteExecutionResult { Success = false };

            try
            {
                if (string.IsNullOrEmpty(sentenceNames))
                {
                    result.ErrorMessage = "缺少 deletion.codes";
                    return result;
                }

                if (!displayStartPosition.HasValue)
                {
                    result.ErrorMessage = "缺少唯一定位结果";
                    return result;
                }

                var names = ParseCodeList(sentenceNames);
                if (names.Count == 0)
                {
                    result.ErrorMessage = "deletion.codes 为空";
                    return result;
                }

                if (names.Any(n => n.StartsWith("T_", StringComparison.Ordinal)
                    || n.StartsWith("C_", StringComparison.Ordinal)))
                {
                    result.ErrorMessage = "delete 仅支持 S_ 编码；T_/C_ 仅用于 insert 锚点";
                    return result;
                }

                result.AffectedCodes = names;

                Word.Range targetRange = DisplayPositionRangeResolver.ResolveSpan(
                    doc,
                    names,
                    displayStartPosition.Value,
                    tableId,
                    tableScope,
                    debugTag: "process_actions_delete");

                if (targetRange == null)
                {
                    result.ErrorMessage = "未找到任何句子 Range";
                    System.Diagnostics.Debug.WriteLine($"[删除操作] 未找到任何句子Range");
                    return result;
                }

                int cursorPos = targetRange.Start;
                targetRange.Delete();

                MoveSelectionAfterDelete(doc, cursorPos);

                result.Success = true;
                System.Diagnostics.Debug.WriteLine($"[删除操作] 已删除 {names.Count} 个句子");
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                System.Diagnostics.Debug.WriteLine($"[删除操作] 删除失败: {ex.Message}");
                return result;
            }
        }

        private static int GetInsertSortPosition(Word.Document doc, ParagraphAction action, TableScopeIndex tableScope)
        {
            if (action?.TargetLocator == null)
            {
                return int.MaxValue;
            }

            if (action.TargetLocator.Kind == LocatorKind.Sentence)
            {
                return GetSentencePositionInDocument(
                    doc,
                    action.TargetLocator.Codes,
                    action.ResolvedTargetDisplayStart,
                    action.TargetLocator.InTableId,
                    tableScope);
            }

            return int.MaxValue - 1;
        }

        /// <summary>
        /// 按 target 定位后插入（仅后插）：句 Collapse(End)；表/图走 ObjectInsertRangeHelper；无 target 则光标处。
        /// </summary>
        private static InsertExecutionResult ExecuteInsert(
            Word.Document doc,
            LocatorSpec target,
            string insertContent,
            int? targetDisplayStart = null,
            TableScopeIndex tableScope = null,
            Dictionary<string, object> formatSnapshot = null,
            string formatMode = null)
        {
            var result = new InsertExecutionResult { Success = false, NewSentences = new List<string>() };

            try
            {
                if (string.IsNullOrEmpty(insertContent))
                {
                    result.ErrorMessage = "缺少 insert 内容";
                    return result;
                }

                Word.Range insertRange = null;

                if (target == null)
                {
                    insertRange = doc.Application.Selection.Range;
                    System.Diagnostics.Debug.WriteLine($"[ProcessActions][insert] 无 target，光标处插入");
                }
                else if (target.Kind == LocatorKind.Table)
                {
                    insertRange = ObjectInsertRangeHelper.FindRangeAfterTable(doc, target.TableId);
                }
                else if (target.Kind == LocatorKind.Chart)
                {
                    insertRange = ObjectInsertRangeHelper.FindRangeAfterChart(doc, target.ChartId);
                }
                else if (target.Kind == LocatorKind.Image)
                {
                    insertRange = ObjectInsertRangeHelper.FindRangeAfterImage(doc, target.ImageId);
                }
                else if (target.Kind == LocatorKind.Sentence)
                {
                    var names = target.Codes.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
                    if (names.Count == 0)
                    {
                        result.ErrorMessage = "target.codes 为空";
                        return result;
                    }

                    if (!targetDisplayStart.HasValue)
                    {
                        result.ErrorMessage = "缺少 target 锚点唯一定位结果";
                        return result;
                    }

                    Word.Range span = DisplayPositionRangeResolver.ResolveSpan(
                        doc,
                        names,
                        targetDisplayStart.Value,
                        target.InTableId,
                        tableScope,
                        debugTag: "process_actions_insert_target");
                    if (span != null)
                    {
                        insertRange = span;
                        insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }

                if (insertRange == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ProcessActions][insert] 无法确定插入位置");
                    result.ErrorMessage = "无法确定插入位置";
                    return result;
                }

                List<string> insertSentences = SentenceSplitter.SplitForActionContent(insertContent);
                if (insertSentences.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[ProcessActions][insert] 插入内容无法切分为句子");
                    result.ErrorMessage = "插入内容无法切分为句子";
                    return result;
                }

                result.NewSentences = insertSentences;

                if (formatSnapshot == null || formatSnapshot.Count == 0 || string.IsNullOrEmpty(formatMode))
                {
                    result.ErrorMessage = ActionFormatHelper.MissingFormatError;
                    return result;
                }

                var insertContentBuilder = new StringBuilder();
                foreach (var sentence in insertSentences)
                {
                    insertContentBuilder.Append(sentence);
                }

                string insertText = insertContentBuilder.ToString();
                insertRange = WordNoteMarkHelper.WriteTextWithNoteMarks(doc, insertRange, insertText)
                    ?? insertRange;
                MoveSelectionAfterInsert(doc, insertRange);

                try
                {
                    Word.Range insertedBlock = doc.Range(insertRange.Start, insertRange.End);
                    List<string> sentencesForFormat = insertSentences.ConvertAll(TextUtils.TrimTrailingParagraphMarks);
                    FormatInheritHelper.ApplySnapshotToSentences(
                        doc, insertedBlock, sentencesForFormat, formatSnapshot, FindSentenceInRange);
                    result.FormatMode = formatMode;
                    FormatInheritHelper.DbgLog(
                        $"action=insert format_mode={formatMode} sentences={insertSentences.Count}");
                }
                catch (Exception fmtEx)
                {
                    result.FormatApplyError = fmtEx.Message;
                    result.ErrorMessage = $"文字已写入，但字符格式套用失败: {fmtEx.Message}";
                    System.Diagnostics.Debug.WriteLine($"[格式] insert 套用失败: {fmtEx.Message}");
                    return result;
                }

                result.Success = true;
                System.Diagnostics.Debug.WriteLine($"[ProcessActions][insert] 已插入 {insertSentences.Count} 个句子");
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                System.Diagnostics.Debug.WriteLine($"[ProcessActions][insert] 插入失败: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// 在指定的 Range 内查找句子
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="searchRange">搜索范围</param>
        /// <param name="sentenceText">句子文本</param>
        /// <returns>找到的 Range，如果未找到返回 null</returns>
        private static Word.Range FindSentenceInRange(Word.Document doc, Word.Range searchRange, string sentenceText)
        {
            try
            {
                if (doc == null || searchRange == null || string.IsNullOrEmpty(sentenceText))
                {
                    return null;
                }
                
                // 记录搜索范围的边界
                int searchStart = searchRange.Start;
                int searchEnd = searchRange.End;
                int searchLength = searchEnd - searchStart;
                
                // 转换换行符为Word代码格式
                string convertedText = ConvertNewlinesToWordCodes(sentenceText);
                string sentenceEscaped = sentenceText.Replace("\r", "\\r").Replace("\n", "\\n");
                string convertedEscaped = convertedText.Replace("^p", "\\^p").Replace("^l", "\\^l");
                
                FormatInheritHelper.DbgLog(
                    $"[FindSentenceInRange] 搜索范围: Start={searchStart}, End={searchEnd}, 长度={searchLength}, 句子长度={sentenceText.Length}, 句子(转义)=\"{sentenceEscaped}\", 转换后(转义)=\"{convertedEscaped}\"");
                
                // 创建新的 Range 用于搜索，避免修改原始 searchRange
                Word.Range findRange = doc.Range(searchStart, searchEnd);
                findRange.Find.ClearFormatting();
                findRange.Find.Text = convertedText;
                findRange.Find.MatchCase = true;
                findRange.Find.MatchWholeWord = false;
                findRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                
                bool executeResult = findRange.Find.Execute();
                FormatInheritHelper.DbgLog($"[FindSentenceInRange] Find.Execute() = {executeResult}");
                
                if (executeResult)
                {
                    // 确保找到的 range 在搜索范围内
                    int foundStart = findRange.Start;
                    int foundEnd = findRange.End;
                    bool inRange = foundStart >= searchStart && foundEnd <= searchEnd;
                    FormatInheritHelper.DbgLog(
                        $"[FindSentenceInRange] 找到: Start={foundStart}, End={foundEnd}, 在范围内={inRange} (需 searchStart<={foundStart}, foundEnd<={searchEnd})");
                    
                    if (inRange)
                    {
                        // 创建新的 Range 对象返回
                        Word.Range foundRange = doc.Range(foundStart, foundEnd);
                        return foundRange;
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FindSentenceInRange] ❌ 查找句子失败: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// 将换行符转换为Word Find功能的特殊字符代码
        /// </summary>
        /// <param name="text">原始文本</param>
        /// <returns>转换后的文本（\r -> ^p, \n -> ^l, \t -> ^t）</returns>
        private static string ConvertNewlinesToWordCodes(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            string result = text.Replace("\r\n", "^p");
            result = result.Replace("\r", "^p");
            result = result.Replace("\n", "^l");
            result = result.Replace("\t", "^t");
            
            return result;
        }

        private static string ValidateSentenceSequenceContinuity(
            string sentenceNames,
            int? displayStartPosition,
            string actionType,
            string tableId = null,
            TableScopeIndex tableScope = null)
        {
            if (string.IsNullOrEmpty(sentenceNames))
            {
                return null;
            }

            var codeList = ParseCodeList(sentenceNames);
            if (codeList.Count == 0)
            {
                return null;
            }

            return ValidateSentenceSequenceContinuityCore(
                codeList,
                displayStartPosition,
                actionType,
                tableId,
                tableScope);
        }

        /// <summary>
        /// 验证句子序列在快照中是否连续（兼容旧调用，默认 displayStartPosition=0）
        /// </summary>
        private static string ValidateSentenceSequenceContinuity(string sentenceNames, string actionType)
        {
            return ValidateSentenceSequenceContinuity(sentenceNames, 0, actionType);
        }

        private static string ValidateSentenceSequenceContinuityCore(
            List<string> codeList,
            int? displayStartPosition,
            string actionType,
            string tableId = null,
            TableScopeIndex tableScope = null)
        {
            // 判断编码类型（检查第一个编码的前缀）
            bool isTableCode = codeList[0].StartsWith("T_");
            bool isSentenceCode = codeList[0].StartsWith("S_");
            bool isChartCode = codeList[0].StartsWith("C_");

            // 验证所有编码类型一致
            foreach (var code in codeList)
            {
                if (isTableCode && !code.StartsWith("T_"))
                {
                    return $"错误：编码序列中混合了不同类型的编码。表格编码（T_）、图表编码（C_）和句子编码（S_）不能混用。";
                }
                if (isChartCode && !code.StartsWith("C_"))
                {
                    return $"错误：编码序列中混合了不同类型的编码。表格编码（T_）、图表编码（C_）和句子编码（S_）不能混用。";
                }
                if (isSentenceCode && !code.StartsWith("S_"))
                {
                    return $"错误：编码序列中混合了不同类型的编码。表格编码（T_）、图表编码（C_）和句子编码（S_）不能混用。";
                }
            }

            List<string> snapshotList;
            string snapshotType;

            if (isTableCode)
            {
                // T_编码：检查表格顺序列表
                snapshotList = DocumentState.GetTableIdOrder();
                snapshotType = "表格";
            }
            else if (isChartCode)
            {
                // C_编码：检查图表顺序列表
                snapshotList = DocumentState.GetChartIdOrder();
                snapshotType = "图表";
            }
            else if (isSentenceCode)
            {
                if (!string.IsNullOrEmpty(tableId) && tableScope != null)
                {
                    if (!tableScope.TryGetTableSentenceCodes(tableId, out IReadOnlyList<string> tableCodes)
                        || tableCodes == null
                        || tableCodes.Count == 0)
                    {
                        return $"错误：表格 {tableId} 内无句子编码，无法验证表内连续性。";
                    }

                    snapshotList = tableCodes.ToList();
                    snapshotType = $"表格 {tableId} 内句子";
                }
                else
                {
                    string snapshot = DocumentState.Snapshot;
                    if (string.IsNullOrEmpty(snapshot))
                    {
                        return $"错误：无法验证编码序列连续性，因为快照为空。请先调用F_get_document_content工具获取文档内容。";
                    }

                    snapshotList = snapshot.Split(',')
                        .Select(s => s.Trim())
                        .Where(s => !string.IsNullOrEmpty(s))
                        .ToList();
                    snapshotType = "句子";
                }
            }
            else
            {
                // 兼容旧格式：默认按句子处理
                string snapshot = DocumentState.Snapshot;
                if (string.IsNullOrEmpty(snapshot))
                {
                    return $"错误：无法验证编码序列连续性，因为快照为空。请先调用F_get_document_content工具获取文档内容。";
                }

                snapshotList = snapshot.Split(',')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToList();
                snapshotType = "句子";
            }

            if (snapshotList.Count == 0)
            {
                return $"错误：无法验证编码序列连续性，因为快照为空。请先调用F_get_document_content工具获取文档内容。";
            }

            // 在快照中查找每个编码的位置（消歧后用 display_start_position）
            var positions = new List<(string name, int index)>();
            for (int i = 0; i < codeList.Count; i++)
            {
                string codeName = codeList[i];
                if (isSentenceCode || (!isTableCode && !isSentenceCode && !isChartCode))
                {
                    if (!displayStartPosition.HasValue)
                    {
                        return $"错误：{snapshotType}编码序列缺少唯一定位结果，无法验证连续性。";
                    }

                    int position = displayStartPosition.Value + i;
                    if (!string.IsNullOrEmpty(tableId) && tableScope != null)
                    {
                        if (!tableScope.TryGetTableSentenceCodes(tableId, out IReadOnlyList<string> tableCodes)
                            || tableCodes == null
                            || position >= tableCodes.Count
                            || !string.Equals(tableCodes[position], codeName, StringComparison.Ordinal))
                        {
                            return $"错误：{snapshotType}编码 \"{codeName}\" 在表格 {tableId} 内不连续或不存在。";
                        }
                    }
                    else
                    {
                        var snapshotCodes = DocumentState.GetOrderedSentenceCodesFromSnapshot(DocumentState.Snapshot);
                        if (position >= snapshotCodes.Count
                            || !string.Equals(snapshotCodes[position], codeName, StringComparison.Ordinal))
                        {
                            return $"错误：{snapshotType}编码 \"{codeName}\" 在快照中不连续或不存在。";
                        }
                    }

                    positions.Add((codeName, position));
                    continue;
                }

                int tableIndex = snapshotList.IndexOf(codeName);
                if (tableIndex < 0)
                {
                    return $"错误：{snapshotType}编码 \"{codeName}\" 在快照中不存在。请先调用F_get_document_content工具获取当前文档的编码序列。";
                }

                positions.Add((codeName, tableIndex));
            }

            // 按位置排序
            positions = positions.OrderBy(p => p.index).ToList();

            // 如果只有一个句子，总是连续的
            if (positions.Count <= 1)
                return null;

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
            // 格式："S_00000","S_00001" 和 "S_00002","S_00003" 或 "T_eStyRRUe" 和 "T_abc123"
            var groupStrings = groups.Select(g => "\"" + string.Join("\",\"", g) + "\"").ToList();
            var groupDetails = string.Join(" 和 ", groupStrings);

            // 构造建议的正确格式
            var suggestedActions = new List<string>();
            for (int i = 0; i < groups.Count; i++)
            {
                    var groupSentenceNames = string.Join(",", groups[i]);
                    if (actionType.ToLower().StartsWith("replace"))
                    {
                        suggestedActions.Add($"{{\"type\": \"replace\",\"detail\": {{\"original\": {{\"codes\": \"{groupSentenceNames}\"}},\"new\": \"新内容{i + 1}\"}}}}");
                    }
                    else if (actionType.ToLower().StartsWith("delete"))
                    {
                        suggestedActions.Add($"{{\"type\": \"delete\",\"detail\": {{\"deletion\": {{\"codes\": \"{groupSentenceNames}\"}}}}}}");
                    }
                    else if (actionType.ToLower().StartsWith("insert"))
                    {
                        suggestedActions.Add($"{{\"type\": \"insert\",\"detail\": {{\"target\": {{\"codes\": \"{groupSentenceNames}\"}},\"insert\": \"插入内容{i + 1}\"}}}}");
                    }
            }

            var suggestedFormat = "{\"action\": [" + string.Join(",", suggestedActions) + "]}";

            return $"{groupDetails}不连续，需要分开来写，应该的结构是：{suggestedFormat}";
        }

        /// <summary>
        /// 删除文本中的句子编号标识，如 [S_00000-start] 和 [S_00000-end]
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <returns>删除句子编号标识后的文本</returns>
        private static string RemoveSentenceIdentifiers(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            
            // 使用正则表达式删除句子编号标识
            // 格式：方括号内S_前缀加5位编码（数字、小写字母、大写字母）加-start或-end，如 [S_00000-start] 和 [S_00000-end]
            // 正则表达式：\[S_[0-9a-zA-Z]{5}-start\]|\[S_[0-9a-zA-Z]{5}-end\]
            string pattern = @"\[S_[0-9a-zA-Z]{5}-start\]|\[S_[0-9a-zA-Z]{5}-end\]";
            string result = Regex.Replace(text, pattern, "");
            
            return result;
        }

        /// <summary>
        /// 段落操作类
        /// </summary>
        private class ParagraphAction
        {
            public int ActionIndex { get; set; }
            public string Type { get; set; }
            public int TargetIndex { get; set; } // 句子序号（不再使用）
            public string Position { get; set; }
            public string Content { get; set; }
            
            // 用于通过句子编号定位的字段
            public string SentenceNames { get; set; }
            public LocatorSpec TargetLocator { get; set; }

            public CodeAnchorScopeFields OriginalScopeFields { get; set; } = new CodeAnchorScopeFields();
            public CodeAnchorScopeFields DeletionScopeFields { get; set; } = new CodeAnchorScopeFields();

            public int? ResolvedSentenceDisplayStart { get; set; }
            public int? ResolvedTargetDisplayStart { get; set; }

            /// <summary>replace/insert 的 detail.format；delete 为 null。</summary>
            public ActionFormatHelper.Spec FormatSpec { get; set; }
        }

        private class CodeAnchorScopeFields
        {
            public string TableId { get; set; }
        }

        private static List<string> ParseCodeList(string codes)
        {
            if (string.IsNullOrEmpty(codes))
            {
                return new List<string>();
            }

            return codes.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
        }

        private static void ApplySnapshotDiffToResults(List<ActionResult> modifyResults)
        {
            int postIndex = DocumentState.LatestSnapshotHistoryIndex;
            SnapshotRecord postRecord = DocumentState.GetSnapshotAt(postIndex);
            SnapshotRecord preRecord = FindLatestStep1Before(postIndex);

            foreach (ActionResult mr in modifyResults)
            {
                if (!mr.success || (mr.type != "replace" && mr.type != "insert"))
                {
                    continue;
                }

                SnapshotDiffResult diff = null;
                if (mr.type == "replace" && mr.affected_codes != null && mr.affected_codes.Count > 0)
                {
                    diff = SnapshotDiffHelper.BuildForReplace(preRecord, postRecord, mr.affected_codes);
                }
                else if (mr.type == "insert")
                {
                    diff = SnapshotDiffHelper.BuildForInsert(
                        preRecord,
                        postRecord,
                        mr.InsertAnchorCodes,
                        null);
                }

                if (diff != null)
                {
                    ApplySnapshotDiffToActionResult(mr, diff);
                }

                if (mr.type == "replace"
                    && mr.PendingPreReplaceSnapshot != null
                    && mr.new_codes != null
                    && mr.new_codes.Count > 0)
                {
                    DocumentState.AttachPreReplaceToNewCodes(mr.new_codes, mr.PendingPreReplaceSnapshot);
                    FormatContextHelper.DbgLog(
                        $"Attach pre_replace → new_codes=[{string.Join(",", mr.new_codes)}]");
                }

                LogPendingVsSnapshotDiff(mr);
            }
        }

        private static SnapshotRecord FindLatestStep1Before(int postIndex)
        {
            for (int i = DocumentState.SnapshotHistory.Count - 1; i >= 0; i--)
            {
                SnapshotRecord record = DocumentState.SnapshotHistory[i];
                if (record.Index < postIndex
                    && string.Equals(record.Source, "process_actions_step1", StringComparison.Ordinal))
                {
                    return record;
                }
            }

            return null;
        }

        private static void ApplySnapshotDiffToActionResult(ActionResult mr, SnapshotDiffResult diff)
        {
            mr.new_sentences = diff.Sentences
                .Select(s => new Dictionary<string, object>
                {
                    ["code"] = s.Code,
                    ["text"] = s.Text ?? string.Empty
                })
                .ToList();
            mr.new_codes = diff.IsVerified
                ? diff.Sentences.Select(s => s.Code).ToList()
                : new List<string>();
            mr.new_codes_source = diff.Source;
            mr.new_codes_warning = diff.Warning;

            if (!diff.IsVerified && mr.success)
            {
                string op = mr.type == "insert" ? "插入" : "替换";
                mr.message =
                    $"{op}成功，但未能可靠解析新句子编码，请调用 F_get_document_content。";
            }
        }

        private static void LogPendingVsSnapshotDiff(ActionResult mr)
        {
            if (mr.PendingNewSentences == null || mr.PendingNewSentences.Count == 0)
            {
                return;
            }

            string snapshotCodes = mr.new_codes != null ? string.Join(",", mr.new_codes) : string.Empty;
            System.Diagnostics.Debug.WriteLine(
                $"[new_sentences] pending_sentence_count={mr.PendingNewSentences.Count} " +
                $"snapshot_codes={snapshotCodes} source={mr.new_codes_source}");
        }

        private static bool ShouldUseBulkSnapshotResults(List<ActionResult> orderedResults)
        {
            int codeableCount = orderedResults.Count(r =>
                r.success && (r.type == "replace" || r.type == "insert"));
            if (codeableCount < 2)
            {
                return false;
            }

            return orderedResults.Any(r =>
                r.success
                && (r.type == "replace" || r.type == "insert")
                && string.Equals(r.new_codes_source, SnapshotDiffHelper.SourceUnavailable, StringComparison.Ordinal));
        }

        private static Dictionary<string, object> BuildBulkSnapshotResultsObject(List<ActionResult> orderedResults)
        {
            int postIndex = DocumentState.LatestSnapshotHistoryIndex;
            SnapshotRecord postRecord = DocumentState.GetSnapshotAt(postIndex);
            SnapshotRecord preRecord = FindLatestStep1Before(postIndex);

            var affectedUnion = new List<string>();
            var bookmarkPairs = new List<SnapshotDiffHelper.BookmarkPair>();

            foreach (ActionResult mr in orderedResults)
            {
                if (!mr.success || (mr.type != "replace" && mr.type != "insert"))
                {
                    continue;
                }

                foreach (string code in mr.affected_codes ?? new List<string>())
                {
                    if (!affectedUnion.Contains(code))
                    {
                        affectedUnion.Add(code);
                    }
                }

                if (preRecord == null)
                {
                    continue;
                }

                if (mr.type == "replace"
                    && SnapshotDiffHelper.TryDeriveReplaceBookmarks(
                        preRecord,
                        mr.affected_codes,
                        out string left,
                        out string right,
                        out _))
                {
                    bookmarkPairs.Add(new SnapshotDiffHelper.BookmarkPair
                    {
                        LeftBookmark = left,
                        RightBookmark = right
                    });
                }
                else if (mr.type == "insert"
                         && SnapshotDiffHelper.TryDeriveInsertBookmarks(
                             preRecord,
                             mr.InsertAnchorCodes,
                             null,
                             out string insertLeft,
                             out string insertRight,
                             out _))
                {
                    bookmarkPairs.Add(new SnapshotDiffHelper.BookmarkPair
                    {
                        LeftBookmark = insertLeft,
                        RightBookmark = insertRight
                    });
                }
            }

            if (bookmarkPairs.Count == 0)
            {
                return new Dictionary<string, object>
                {
                    ["affected_codes"] = affectedUnion,
                    ["new_sentences"] = new List<Dictionary<string, object>>(),
                    ["new_codes_source"] = SnapshotDiffHelper.SourceUnavailable,
                    ["new_codes_warning"] = "ambiguous_interval",
                    ["message"] = "本批操作已完成，但未能可靠解析新句子编码，请调用 F_get_document_content。"
                };
            }

            SnapshotDiffResult bulk = SnapshotDiffHelper.BuildBulkEnvelope(preRecord, postRecord, bookmarkPairs);
            var dict = new Dictionary<string, object>
            {
                ["affected_codes"] = affectedUnion,
                ["new_sentences"] = bulk.Sentences
                    .Select(s => new Dictionary<string, object>
                    {
                        ["code"] = s.Code,
                        ["text"] = s.Text ?? string.Empty
                    })
                    .ToList(),
                ["new_codes_source"] = bulk.Source
            };

            if (!string.IsNullOrEmpty(bulk.Warning))
            {
                dict["new_codes_warning"] = bulk.Warning;
            }

            if (bulk.IsVerified && bulk.Sentences.Count > 0)
            {
                dict["new_codes"] = bulk.Sentences.Select(s => s.Code).ToList();
            }

            if (!bulk.IsVerified)
            {
                dict["message"] =
                    "本批操作已完成，但未能可靠解析新句子编码，请调用 F_get_document_content。";
            }

            return dict;
        }

        private static void FillReplaceActionResult(ActionResult ar, ReplaceExecutionResult rep)
        {
            ar.affected_codes = rep.OriginalCodes ?? new List<string>();
            ar.PendingNewSentences = rep.NewSentences ?? new List<string>();
            ar.PendingPreReplaceSnapshot = rep.PendingPreReplaceSnapshot;

            if (!rep.Success)
            {
                ar.success = false;
                ar.message = rep.ErrorMessage ?? "替换失败";
                return;
            }

            ar.success = true;
            ar.format_mode = rep.FormatMode;
            ar.format_apply_error = rep.FormatApplyError;

            if (rep.FormatMode == "inherit_from")
            {
                ar.format_applied = "inherit_from";
                ar.format_source = "detail.format.inherit_from";
            }
            else if (rep.FormatMode == "explicit")
            {
                ar.format_applied = "char_format";
                ar.format_source = "detail.format.char_format";
            }

            ar.message = "替换成功";
        }

        private static void FillDeleteActionResult(ActionResult ar, DeleteExecutionResult del)
        {
            ar.affected_codes = del.AffectedCodes ?? new List<string>();
            ar.success = del.Success;
            ar.message = del.Success ? "删除成功" : (del.ErrorMessage ?? "删除失败");
        }

        private static void FillInsertActionResult(ActionResult ar, InsertExecutionResult ins, LocatorSpec target)
        {
            string anchorCodes = null;
            if (target != null)
            {
                if (target.Kind == LocatorKind.Sentence)
                {
                    anchorCodes = target.Codes;
                }
                else if (target.Kind == LocatorKind.Table)
                {
                    anchorCodes = target.TableId;
                }
                else if (target.Kind == LocatorKind.Chart)
                {
                    anchorCodes = target.ChartId;
                }
                else if (target.Kind == LocatorKind.Image)
                {
                    anchorCodes = target.ImageId;
                }
            }

            ar.affected_codes = ParseCodeList(anchorCodes);
            ar.PendingNewSentences = ins.NewSentences ?? new List<string>();
            ar.InsertAnchorCodes = ParseCodeList(anchorCodes);
            ar.success = ins.Success;
            ar.format_mode = ins.FormatMode;
            ar.format_apply_error = ins.FormatApplyError;
            if (ins.FormatMode == "inherit_from")
            {
                ar.format_applied = "inherit_from";
                ar.format_source = "detail.format.inherit_from";
            }
            else if (ins.FormatMode == "explicit")
            {
                ar.format_applied = "char_format";
                ar.format_source = "detail.format.char_format";
            }

            ar.message = ins.Success ? "插入成功" : (ins.ErrorMessage ?? "插入失败");
        }

        private sealed class FormatInheritContext
        {
            public Dictionary<string, object> Snapshot { get; set; }
            public bool FormatWarning { get; set; }
            public string FingerprintFirst { get; set; }
            public string FingerprintSecond { get; set; }
            public Dictionary<string, string> ReplacedFingerprints { get; set; }
        }

        private sealed class ReplaceExecutionResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public List<string> OriginalCodes { get; set; }
            public List<string> NewSentences { get; set; }
            public string FormatMode { get; set; }
            public string FormatApplyError { get; set; }
            public Dictionary<string, object> PendingPreReplaceSnapshot { get; set; }
        }

        private sealed class DeleteExecutionResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public List<string> AffectedCodes { get; set; }
        }

        private sealed class InsertExecutionResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public List<string> NewSentences { get; set; }
            public string FormatMode { get; set; }
            public string FormatApplyError { get; set; }
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
            public List<string> affected_codes { get; set; } = new List<string>();
            public List<string> new_codes { get; set; } = new List<string>();
            public List<Dictionary<string, object>> new_sentences { get; set; } = new List<Dictionary<string, object>>();
            public string new_codes_source { get; set; }
            public string new_codes_warning { get; set; }
            public List<string> InsertAnchorCodes { get; set; } = new List<string>();
            public string format_mode { get; set; }
            public string format_applied { get; set; }
            public string format_source { get; set; }
            public string format_apply_error { get; set; }
            public List<string> PendingNewSentences { get; set; }
            public Dictionary<string, object> PendingPreReplaceSnapshot { get; set; }

            public object ToResponseObject()
            {
                var dict = new Dictionary<string, object>
                {
                    ["action_index"] = action_index,
                    ["type"] = type,
                    ["success"] = success,
                    ["message"] = message ?? ""
                };

                if (affected_codes != null && affected_codes.Count > 0)
                {
                    dict["affected_codes"] = affected_codes;
                }

                if (type == "replace" || type == "insert")
                {
                    dict["new_sentences"] = new_sentences ?? new List<Dictionary<string, object>>();
                }

                if (new_codes != null && new_codes.Count > 0)
                {
                    dict["new_codes"] = new_codes;
                }

                if (!string.IsNullOrEmpty(new_codes_source))
                {
                    dict["new_codes_source"] = new_codes_source;
                }

                if (!string.IsNullOrEmpty(new_codes_warning))
                {
                    dict["new_codes_warning"] = new_codes_warning;
                }

                if (!string.IsNullOrEmpty(format_mode))
                {
                    dict["format_mode"] = format_mode;
                }

                if (!string.IsNullOrEmpty(format_applied))
                {
                    dict["format_applied"] = format_applied;
                }

                if (!string.IsNullOrEmpty(format_source))
                {
                    dict["format_source"] = format_source;
                }

                if (!string.IsNullOrEmpty(format_apply_error))
                {
                    dict["format_apply_error"] = format_apply_error;
                }

                return dict;
            }
        }
    }
}

