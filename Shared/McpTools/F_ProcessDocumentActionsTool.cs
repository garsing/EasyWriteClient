using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using WordAddIn1.DocumentMapping;
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
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tool-enter",
                        WordAddIn1.DocumentHost.HostProbe.ArgsBrief(args));
                    if (TryGetContractParseFailure(args, out ToolResult parseFail))
                    {
                        WordAddIn1.DocumentHost.HostProbe.Log(
                            "tool-skip-resolve-parse",
                            parseFail == null ? "" : parseFail.Error);
                        return parseFail;
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        WordAddIn1.DocumentHost.HostProbe.Log(
                            "tool-resolve-fail",
                            resolveError == null ? "" : resolveError.Error);
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_process_document_actions] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return resolveError;
                    }

                    Word.Document document = docHandle.Document;
                    WordAddIn1.DocumentHost.HostProbe.Touch("tool-doc", document);

                    var swTotal = Stopwatch.StartNew();
                    int timedActionCount = 0;
                    try
                    {

                    string logPath = EasyWriteLog.BeginSession("process_document_actions");
                    WrittenFormatGlueHelper.ClearPlans();
                    DocumentActionPathLog.Clear();
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 本次测试日志: {logPath}");
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_process_document_actions] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作前文档状态: 段落总数={document.Paragraphs.Count}");
                    FormatInheritHelper.DbgLogDocumentFontSample(document, "process_actions 入口");

                    FormatInheritHelper.DbgLog("门面整批预检前");
                    FormatInheritHelper.DbgLogDocumentFontSample(document, "门面预检前");

                    System.Diagnostics.Debug.WriteLine("[DEBUG] 开始解析操作参数");
                    (bool success, string error, List<ParagraphAction> actions) parseResult;
                    using (EasyWriteDiagnostics.Time("process_actions.parse"))
                    {
                        parseResult = ParseActions(args["action"], document);
                    }

                    if (!parseResult.success)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ERROR] 解析操作失败: {parseResult.error}");
                        return FailBeforeWrite(parseResult.error);
                    }

                    List<ParagraphAction> actions = parseResult.actions;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 解析操作成功，共 {actions.Count} 个操作");
                    timedActionCount = actions.Count;

                    string unusable = FindInoperableActionError(actions);
                    if (!string.IsNullOrEmpty(unusable))
                    {
                        return FailBeforeWrite(unusable);
                    }

                    for (int i = 0; i < actions.Count; i++)
                    {
                        actions[i].ActionIndex = i;
                    }

                    InsertTargetLocatorHelper.PrecheckResult precheck;
                    using (EasyWriteDiagnostics.Time("process_actions.precheck"))
                    {
                        precheck = InsertTargetLocatorHelper.TryPrecheckTargets(
                            document,
                            BuildAmbiguityAnchors(actions),
                            docHandle.DisallowBackendApi);
                    }
                    if (!precheck.Success)
                    {
                        return InsertTargetLocatorHelper.ToPrecheckFailureResult(
                            precheck,
                            document.Name ?? "未命名文档",
                            actions.Count);
                    }

                    TableScopeIndex tableScope = precheck.TableScope
                        ?? TableSentenceScopeHelper.BuildIndexFromProcessCache(DocumentState.LastProcessCache);
                    ApplyResolvedDisplayPositions(actions, precheck.Ambiguity);

                    string continuityError;
                    using (EasyWriteDiagnostics.Time("process_actions.continuity"))
                    {
                        continuityError = ValidateAllActionsContinuity(actions, tableScope);
                    }

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
                    List<ActionResult> modifyResults;
                    using (EasyWriteDiagnostics.Time("process_actions.modify"))
                    {
                        modifyResults = await ModifyDocumentDirectly(
                            document,
                            actions,
                            tableScope,
                            docHandle.DisallowBackendApi);
                    }
                    FormatInheritHelper.DbgLogDocumentFontSample(document, "第三步 modify 后");
                    System.Diagnostics.Debug.WriteLine("[第三步完成]");

                    var orderedResults = modifyResults.OrderBy(r => r.action_index).ToList();
                    object resultsPayload = orderedResults.Select(r => r.ToResponseObject()).ToList();
                    string batchMessage = null;
                    bool bulkMode = false;
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
                    List<Dictionary<string, object>> navigates;
                    Dictionary<string, object> navigateSummary;
                    using (EasyWriteDiagnostics.Time("process_actions.navigate"))
                    {
                        var navigateOutcome = DocumentNavigateHelper.BuildAndApply(
                            wordApplicationTyped,
                            activeDocument,
                            navigateInputs,
                            bulkMode,
                            navigateSettings.AutoNavigateAfterModify);
                        navigates = navigateOutcome.Navigates;
                        navigateSummary = navigateOutcome.Summary;
                    }
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
                            navigate = navigateSummary,
                            format_apply_log = WrittenFormatGlueHelper.SnapshotPlans(),
                            geo_plan_log = DocumentActionPathLog.Snapshot()
                        }
                    };

                    return result;
                    }
                    finally
                    {
                        EasyWriteDiagnostics.LogTiming(
                            "process_actions.total",
                            swTotal.ElapsedMilliseconds,
                            "actions=" + timedActionCount);
                    }
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
        /// 不读稿的契约解析。失败则调用方不应建副本、不应 resolve。
        /// </summary>
        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            if (args == null || !args.ContainsKey("action"))
            {
                failure = FailBeforeWrite("缺少必需参数：action");
                return true;
            }

            var parseResult = ParseActions(args["action"], null);
            if (!parseResult.Item1)
            {
                failure = FailBeforeWrite(parseResult.Item2);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 尚未写字的契约失败：只丢副本，不按 Word.ActiveDocument 回滚正文。
        /// </summary>
        private static ToolResult FailBeforeWrite(string error)
        {
            return new ToolResult
            {
                Success = false,
                Error = error,
                Data = new Dictionary<string, object>
                {
                    ["checkpoint_cleanup_only"] = true
                }
            };
        }

        /// <summary>
        /// 解析操作列表：replace/delete/insert；定位走统一定位门面（含 paragraph_codes）。
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
                        var actionDict = ActionFormatHelper.CoerceDict(actionObj);
                        if (actionDict == null || !actionDict.ContainsKey("type"))
                            continue;

                        var action = new ParagraphAction();
                        action.Type = actionDict["type"]?.ToString();

                        var detail = actionDict.ContainsKey("detail")
                            ? ActionFormatHelper.CoerceDict(actionDict["detail"])
                            : null;
                        if (detail == null)
                            continue;

                        switch (action.Type?.ToLower())
                        {
                            case "replace":
                                object original = detail.ContainsKey("original") ? detail["original"] : null;
                                Dictionary<string, object> originalDict = ActionFormatHelper.CoerceDict(original);
                                if (originalDict == null || originalDict.Count == 0)
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
                                
                                if (!detail.ContainsKey("new"))
                                {
                                    return (false, "replace 缺少 new", null);
                                }

                                action.ContentPayload = ActionContentHelper.TryParseInsertOrNew(detail["new"], detail);
                                if (!action.ContentPayload.Ok)
                                {
                                    return (false, action.ContentPayload.Error ?? "replace.new 无效", null);
                                }

                                action.Content = action.ContentPayload.JoinedText;
                                action.FormatSpec = action.ContentPayload.IsRuns
                                    ? null
                                    : action.ContentPayload.FormatSpec;

                                break;

                            case "delete":
                                object deletion = detail.ContainsKey("deletion") ? detail["deletion"] : null;
                                Dictionary<string, object> deletionDict = ActionFormatHelper.CoerceDict(deletion);
                                if (deletionDict == null || deletionDict.Count == 0)
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
                                if (!detail.ContainsKey("insert"))
                                {
                                    return (false, "insert 缺少 insert 文本", null);
                                }

                                action.ContentPayload = ActionContentHelper.TryParseInsertOrNew(detail["insert"], detail);
                                if (!action.ContentPayload.Ok)
                                {
                                    return (false, action.ContentPayload.Error ?? "insert 文本无效", null);
                                }

                                action.Content = action.ContentPayload.JoinedText;
                                action.FormatSpec = action.ContentPayload.IsRuns
                                    ? null
                                    : action.ContentPayload.FormatSpec;

                                if (detail.ContainsKey("target") && detail["target"] != null)
                                {
                                    Dictionary<string, object> targetDict = ActionFormatHelper.CoerceDict(detail["target"]);
                                    if (targetDict == null || targetDict.Count == 0)
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
            Image,
            DocumentStart,
            EmptyCell,
            Paragraph
        }

        private sealed class LocatorSpec
        {
            public LocatorKind Kind { get; set; }
            public string Codes { get; set; }
            public string InTableId { get; set; }
            public string TableId { get; set; }
            public string ChartId { get; set; }
            public string ImageId { get; set; }
            public int Row { get; set; }
            public int Col { get; set; }
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
                case InsertTargetLocatorHelper.LocatorKind.DocumentStart:
                    kind = LocatorKind.DocumentStart;
                    break;
                case InsertTargetLocatorHelper.LocatorKind.EmptyCell:
                    kind = LocatorKind.EmptyCell;
                    break;
                case InsertTargetLocatorHelper.LocatorKind.Paragraph:
                    kind = LocatorKind.Paragraph;
                    break;
                case InsertTargetLocatorHelper.LocatorKind.Sentence:
                    kind = LocatorKind.Sentence;
                    break;
                default:
                    return new LocatorParseResult
                    {
                        Success = false,
                        Error = $"{fieldLabel} 无法识别的定位类型"
                    };
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
                    ImageId = s.ImageId,
                    Row = s.Row,
                    Col = s.Col
                }
            };
        }

        private static string FindInoperableActionError(List<ParagraphAction> actions)
        {
            if (actions == null)
            {
                return null;
            }

            foreach (ParagraphAction action in actions)
            {
                if (StoryUniqEncode.CsvIsInoperable(action.SentenceNames)
                    || (action.TargetLocator != null
                        && StoryUniqEncode.CsvIsInoperable(action.TargetLocator.Codes)))
                {
                    return StoryUniqEncode.FailBeforeWriteMessage;
                }
            }

            return null;
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
                if (type == "replace" && !string.IsNullOrEmpty(action.SentenceNames)
                    && !EmptyCellCode.IsEmptyCellCode(action.SentenceNames))
                {
                    anchors.Add(ToAmbiguityAnchor(action.ActionIndex, "original", action.SentenceNames, action.OriginalScopeFields));
                }
                else if (type == "delete" && !string.IsNullOrEmpty(action.SentenceNames)
                    && !EmptyCellCode.IsEmptyCellCode(action.SentenceNames))
                {
                    anchors.Add(ToAmbiguityAnchor(action.ActionIndex, "deletion", action.SentenceNames, action.DeletionScopeFields));
                }
                else if (type == "insert"
                         && action.TargetLocator != null
                         && (action.TargetLocator.Kind == LocatorKind.Sentence
                             || action.TargetLocator.Kind == LocatorKind.Paragraph)
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
                if ((type == "replace" || type == "delete")
                    && EmptyCellCode.IsEmptyCellCode(action.SentenceNames))
                {
                    continue;
                }

                if ((type == "replace" || type == "delete")
                    && LooksLikeParagraphCodes(action.SentenceNames))
                {
                    continue;
                }

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
                         && action.TargetLocator.Kind == LocatorKind.EmptyCell)
                {
                    continue;
                }
                else if (type == "insert"
                         && action.TargetLocator != null
                         && action.TargetLocator.Kind == LocatorKind.Paragraph)
                {
                    continue;
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
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            var results = new List<ActionResult>();

            try
            {
                IDisposable sortAll = EasyWriteDiagnostics.Time("process_actions.sort_all");
                var ordered = actions
                    .Select((action, originalIndex) => new IndexedAction
                    {
                        Action = action,
                        OriginalIndex = originalIndex
                    })
                    .OrderByDescending(x => GetActionSortKey(x.Action))
                    .ToList();
                sortAll.Dispose();

                var byIndex = new Dictionary<int, ParagraphAction>();
                foreach (IndexedAction item in ordered)
                {
                    byIndex[item.OriginalIndex] = item.Action;
                }

                List<DocumentActionPlan.Item> planItems = ordered
                    .Select(x => ToPlanItem(x))
                    .ToList();
                List<DocumentActionPlan.Run> plan = DocumentActionPlan.BuildPlan(planItems);
                int geoRuns = 0;
                int singles = 0;
                foreach (DocumentActionPlan.Run step in plan)
                {
                    if (step.IsGeoRun)
                    {
                        geoRuns++;
                    }
                    else
                    {
                        singles++;
                    }
                }

                DocumentActionPathLog.Record(
                    DocumentActionPathLog.LogTag + " plan geo_runs=" + geoRuns + " singles=" + singles);

                foreach (DocumentActionPlan.Run step in plan)
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (step.IsGeoRun)
                    {
                        var members = new List<IndexedAction>();
                        foreach (int idx in step.OriginalIndices)
                        {
                            members.Add(new IndexedAction
                            {
                                OriginalIndex = idx,
                                Action = byIndex[idx]
                            });
                        }

                        List<ActionResult> geoResults = ExecuteGeoRun(
                            doc, members, tableScope, disallowBackendApi);
                        ActionResult anyOk = geoResults.Find(r => r.success);
                        if (anyOk != null)
                        {
                            await RemapAndAssignNewCodes(doc, anyOk, disallowBackendApi, "geo_run");
                        }

                        results.AddRange(geoResults);
                        continue;
                    }

                    int oneIndex = step.OriginalIndices[0];
                    ParagraphAction oneAction = byIndex[oneIndex];
                    DocumentActionPathLog.Record(
                        DocumentActionPathLog.LogTag
                        + " step=single type=" + (oneAction.Type ?? "?")
                        + " geo=" + (TryGetGeoCell(oneAction, out _) ? "true" : "false"));
                    ActionResult one = ExecuteOneAction(
                        doc, oneAction, oneIndex, tableScope, disallowBackendApi);
                    await RemapAndAssignNewCodes(doc, one, disallowBackendApi, "single");
                    results.Add(one);
                }

                results = results.OrderBy(r => r.action_index).ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[直接修改文档] 修改文档失败: {ex.Message}");
            }

            return results;
        }

        private sealed class IndexedAction
        {
            public ParagraphAction Action { get; set; }
            public int OriginalIndex { get; set; }
        }

        private static int GetActionSortKey(ParagraphAction action)
        {
            if (string.Equals(action?.Type, "insert", StringComparison.OrdinalIgnoreCase))
            {
                return GetInsertSnapshotSortPosition(action);
            }

            return GetSnapshotSortPosition(action?.SentenceNames);
        }

        private static DocumentActionPlan.Item ToPlanItem(IndexedAction item)
        {
            bool geo = TryGetGeoCell(item.Action, out EmptyCellCode.Parsed parsed);
            return new DocumentActionPlan.Item
            {
                OriginalIndex = item.OriginalIndex,
                IsGeoCell = geo,
                TableId = geo ? parsed.TableId : null
            };
        }

        private static bool TryGetGeoCell(ParagraphAction action, out EmptyCellCode.Parsed parsed)
        {
            parsed = null;
            string type = action?.Type?.ToLowerInvariant();
            if (type != "insert" && type != "replace" && type != "delete")
            {
                return false;
            }

            if (type == "insert")
            {
                LocatorSpec target = action.TargetLocator;
                if (target == null || target.Kind != LocatorKind.EmptyCell)
                {
                    return false;
                }

                parsed = new EmptyCellCode.Parsed
                {
                    TableId = target.TableId,
                    Row = target.Row,
                    Col = target.Col,
                    Code = target.Codes
                };
                return !string.IsNullOrEmpty(parsed.TableId);
            }

            List<string> names = ParseCodeList(action.SentenceNames);
            return names.Count == 1 && EmptyCellCode.TryParse(names[0], out parsed);
        }

        private static ActionResult ExecuteOneAction(
            Word.Document doc,
            ParagraphAction action,
            int originalIndex,
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            var result = new ActionResult
            {
                action_index = originalIndex,
                type = action.Type,
                target_index = 0
            };

            try
            {
                using (EasyWriteDiagnostics.Time(
                    "process_actions.action." + (action.Type ?? "?") + "[" + originalIndex + "]"))
                {
                    switch (action.Type?.ToLower())
                    {
                        case "replace":
                        {
                            var rep = ExecuteReplaceWithPayload(
                                doc, action, tableScope, disallowBackendApi);
                            FillReplaceActionResult(result, rep);
                            break;
                        }

                        case "delete":
                        {
                            var del = ExecuteDeleteBySentenceNames(
                                doc,
                                action.SentenceNames,
                                action.DeletionScopeFields?.TableId,
                                disallowBackendApi);
                            FillDeleteActionResult(result, del);
                            break;
                        }

                        case "insert":
                        {
                            var ins = ExecuteInsertWithPayload(
                                doc, action, tableScope, disallowBackendApi);
                            FillInsertActionResult(result, ins, action.TargetLocator);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.success = false;
                result.message = $"操作失败: {ex.Message}";
            }

            return result;
        }

        private static List<ActionResult> ExecuteGeoRun(
            Word.Document doc,
            List<IndexedAction> members,
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            var results = new List<ActionResult>();
            if (members == null || members.Count == 0)
            {
                return results;
            }

            EmptyCellCode.Parsed firstGeo;
            if (!TryGetGeoCell(members[0].Action, out firstGeo))
            {
                foreach (IndexedAction member in members)
                {
                    results.Add(FailAction(member, "几何码集合无法解析第一条"));
                }

                return results;
            }

            DocumentActionPathLog.Record(
                DocumentActionPathLog.LogTag
                + " step=geo_run cells=" + members.Count
                + " table=" + firstGeo.TableId);

            var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                doc,
                new InsertTargetLocatorHelper.LocatorSpec
                {
                    Kind = InsertTargetLocatorHelper.LocatorKind.EmptyCell,
                    Codes = firstGeo.Code,
                    TableId = firstGeo.TableId,
                    Row = firstGeo.Row,
                    Col = firstGeo.Col
                },
                InsertTargetLocatorHelper.LocateRangeKind.WriteCell,
                disallowBackendApi,
                refreshDocument: false);
            DocumentActionPathLog.Record(
                DocumentActionPathLog.LogTag + " locate via=geo_run n=1");
            if (!locate.Success)
            {
                string err = locate.Error ?? "无法定位空格";
                foreach (IndexedAction member in members)
                {
                    results.Add(FailAction(member, err));
                }

                return results;
            }

            Word.Table table = TableResolveHelper.TryResolveTableById(doc, firstGeo.TableId, out string tableError);
            if (table == null)
            {
                string err = tableError ?? "找不到表格";
                foreach (IndexedAction member in members)
                {
                    results.Add(FailAction(member, err));
                }

                return results;
            }

            TableRowValuesHelper.PrepareLogicalCellMap(table);
            DocumentActionPathLog.Record(DocumentActionPathLog.LogTag + " map=once");

            Word.Application app = null;
            try
            {
                app = doc.Application;
            }
            catch
            {
                app = null;
            }

            bool prevScreen;
            if (!TableRowValuesHelper.TrySuppressScreenUpdating(app, out prevScreen))
            {
                app = null;
            }

            try
            {
                bool stop = false;
                string stopReason = null;
                for (int i = 0; i < members.Count; i++)
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    IndexedAction member = members[i];
                    if (stop)
                    {
                        results.Add(FailAction(member, stopReason ?? "集合内前一格失败，本格未写"));
                        continue;
                    }

                    EmptyCellCode.Parsed geo;
                    if (!TryGetGeoCell(member.Action, out geo))
                    {
                        ActionResult bad = FailAction(member, "几何码无效");
                        results.Add(bad);
                        stop = true;
                        stopReason = bad.message;
                        continue;
                    }

                    Word.Cell cell = TableRowValuesHelper.TryGetMappedCell(table, geo.Row, geo.Col);
                    ActionResult written = WriteGeoMember(
                        doc, cell, member, geo, tableScope, disallowBackendApi);
                    results.Add(written);
                    if (!written.success)
                    {
                        stop = true;
                        stopReason = "集合内前一格失败，本格未写";
                    }

                    bool last = i == members.Count - 1;
                    int nextRow = -1;
                    if (!last && TryGetGeoCell(members[i + 1].Action, out EmptyCellCode.Parsed nextGeo))
                    {
                        nextRow = nextGeo.Row;
                    }

                    if (last || nextRow != geo.Row)
                    {
                        TableRowValuesHelper.TryFlashScreen(app);
                        DocumentActionPathLog.Record(
                            DocumentActionPathLog.LogTag + " flash row=" + geo.Row);
                    }
                }
            }
            finally
            {
                TableRowValuesHelper.TryRestoreScreenUpdating(app, prevScreen);
            }

            return results;
        }

        private static ActionResult FailAction(IndexedAction member, string message)
        {
            return new ActionResult
            {
                action_index = member.OriginalIndex,
                type = member.Action?.Type,
                target_index = 0,
                success = false,
                message = message
            };
        }

        private static ActionResult WriteGeoMember(
            Word.Document doc,
            Word.Cell cell,
            IndexedAction member,
            EmptyCellCode.Parsed geo,
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            var result = new ActionResult
            {
                action_index = member.OriginalIndex,
                type = member.Action.Type,
                target_index = 0
            };

            try
            {
            if (cell == null)
            {
                result.success = false;
                result.message = TableRowValuesHelper.BuildUnresolvableCellError(geo.Row, geo.Col);
                return result;
            }

            try
            {
                using (EasyWriteDiagnostics.Time(
                    "process_actions.action.geo[" + member.OriginalIndex + "]"))
                {
                    string type = member.Action.Type?.ToLowerInvariant();
                    if (type == "delete")
                    {
                        var del = ExecuteEmptyCellClear(doc, geo, disallowBackendApi, cell);
                        FillDeleteActionResult(result, del);
                        return result;
                    }

                    ActionContentHelper.Payload payload = member.Action.ContentPayload;
                    if (payload == null || !payload.Ok)
                    {
                        result.success = false;
                        result.message = payload?.Error ?? "内容无效";
                        return result;
                    }

                    if (payload.IsRuns)
                    {
                        var written = ExecuteEmptyCellReplaceRuns(
                            doc, geo, payload, tableScope, disallowBackendApi, cell);
                        if (type == "insert")
                        {
                            var ins = new InsertExecutionResult
                            {
                                Success = written.Success,
                                ErrorMessage = written.ErrorMessage,
                                FormatMode = written.FormatMode,
                                FormatApplyError = written.FormatApplyError,
                                NewSentences = written.NewSentences ?? new List<string>()
                            };
                            FillInsertActionResult(result, ins, member.Action.TargetLocator);
                        }
                        else
                        {
                            FillReplaceActionResult(result, written);
                        }

                        return result;
                    }

                    if (!ActionFormatHelper.TryBuildSnapshot(
                        doc, payload.FormatSpec, tableScope, out var snapshot, out var formatMode, out string fmtErr))
                    {
                        result.success = false;
                        result.message = fmtErr;
                        return result;
                    }

                    var textWritten = ExecuteEmptyCellReplace(
                        doc, geo, payload.Text, snapshot, formatMode, disallowBackendApi, cell);
                    if (type == "insert")
                    {
                        var ins = new InsertExecutionResult
                        {
                            Success = textWritten.Success,
                            ErrorMessage = textWritten.ErrorMessage,
                            FormatMode = textWritten.FormatMode,
                            FormatApplyError = textWritten.FormatApplyError,
                            NewSentences = textWritten.NewSentences ?? new List<string>()
                        };
                        FillInsertActionResult(result, ins, member.Action.TargetLocator);
                    }
                    else
                    {
                        FillReplaceActionResult(result, textWritten);
                    }
                }
            }
            catch (Exception ex)
            {
                result.success = false;
                result.message = $"操作失败: {ex.Message}";
            }

            return result;
            }
            finally
            {
                string type = member.Action?.Type?.ToLowerInvariant() ?? "?";
                DocumentActionPathLog.Record(
                    DocumentActionPathLog.LogTag
                    + " write via=map type=" + type
                    + " r=" + geo.Row
                    + " c=" + geo.Col
                    + " ok=" + (result.success ? "true" : "false"));
            }
        }

        /// <summary>
        /// display 上 S_ / P_ / 几何码 start 的出现序。混批从后往前时句段能夹在两格中间（D5）。
        /// </summary>
        private static List<string> GetDisplayAnchorOrder()
        {
            var ordered = new List<string>();
            string display = GetSortDisplayText();
            if (string.IsNullOrEmpty(display))
            {
                return ordered;
            }

            foreach (Match m in Regex.Matches(
                display,
                @"\[(S_[0-9a-zA-Z]{5}|P_[0-9a-zA-Z]+(?:_[0-9]+)?|T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-start"))
            {
                ordered.Add(m.Groups[1].Value);
            }

            return ordered;
        }

        private static string GetSortDisplayText()
        {
            ProcessDocumentCacheEntry cache = DocumentState.LastProcessCache;
            if (cache != null && cache.ChunkInfo != null && cache.ChunkInfo.Count > 0)
            {
                var sb = new System.Text.StringBuilder();
                foreach (WordDocumentExtractor.ChunkInfo chunk in cache.ChunkInfo)
                {
                    if (chunk != null && !string.IsNullOrEmpty(chunk.DisplayContent))
                    {
                        sb.Append(chunk.DisplayContent);
                    }
                }

                if (sb.Length > 0)
                {
                    return sb.ToString();
                }
            }

            IReadOnlyList<string> parts = DocumentState.ChunkParagraphDisplayContents;
            if (parts == null || parts.Count == 0)
            {
                return "";
            }

            return string.Join("", parts);
        }

        private static int IndexInDisplayAnchors(string codes)
        {
            List<string> displayOrder = GetDisplayAnchorOrder();
            List<string> needle = ParseCodeList(codes);
            if (displayOrder.Count == 0 || needle.Count == 0)
            {
                return -1;
            }

            return displayOrder.IndexOf(needle[0]);
        }

        /// <summary>
        /// 用预检快照里的码序排序，不再 COM Find / ProcessDocument。
        /// </summary>
        private static int GetSnapshotSortPosition(string codes)
        {
            if (string.IsNullOrEmpty(codes))
            {
                return 0;
            }

            int displayIndex = IndexInDisplayAnchors(codes);
            if (displayIndex >= 0)
            {
                return displayIndex;
            }

            EmptyCellCode.Parsed emptyParsed;
            if (EmptyCellCode.TryParse(codes.Split(',')[0].Trim(), out emptyParsed))
            {
                return emptyParsed.Row * 1000 + emptyParsed.Col;
            }

            List<string> ordered = DocumentState.GetOrderedSentenceCodesFromSnapshot(DocumentState.Snapshot);
            List<string> needle = ParseCodeList(codes);
            if (ordered == null || ordered.Count == 0 || needle.Count == 0)
            {
                return 0;
            }

            for (int i = 0; i <= ordered.Count - needle.Count; i++)
            {
                bool match = true;
                for (int j = 0; j < needle.Count; j++)
                {
                    if (!string.Equals(ordered[i + j], needle[j], StringComparison.Ordinal))
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                {
                    return i;
                }
            }

            int one = ordered.IndexOf(needle[0]);
            return one >= 0 ? one : 0;
        }

        private static int GetInsertSnapshotSortPosition(ParagraphAction action)
        {
            if (action?.TargetLocator == null)
            {
                return int.MaxValue;
            }

            LocatorSpec target = action.TargetLocator;
            if (target.Kind == LocatorKind.DocumentStart)
            {
                return 0;
            }

            if (target.Kind == LocatorKind.EmptyCell)
            {
                return GetSnapshotSortPosition(target.Codes);
            }

            if (!string.IsNullOrEmpty(target.Codes))
            {
                return GetSnapshotSortPosition(target.Codes);
            }

            return int.MaxValue - 1;
        }

        private static InsertTargetLocatorHelper.LocatorSpec ToFacadeSpec(LocatorSpec target)
        {
            if (target == null)
            {
                return null;
            }

            var spec = new InsertTargetLocatorHelper.LocatorSpec
            {
                Codes = target.Codes,
                InTableId = target.InTableId,
                TableId = target.TableId,
                ChartId = target.ChartId,
                ImageId = target.ImageId,
                Row = target.Row,
                Col = target.Col
            };
            switch (target.Kind)
            {
                case LocatorKind.Table:
                    spec.Kind = InsertTargetLocatorHelper.LocatorKind.Table;
                    break;
                case LocatorKind.Chart:
                    spec.Kind = InsertTargetLocatorHelper.LocatorKind.Chart;
                    break;
                case LocatorKind.Image:
                    spec.Kind = InsertTargetLocatorHelper.LocatorKind.Image;
                    break;
                case LocatorKind.DocumentStart:
                    spec.Kind = InsertTargetLocatorHelper.LocatorKind.DocumentStart;
                    break;
                case LocatorKind.EmptyCell:
                    spec.Kind = InsertTargetLocatorHelper.LocatorKind.EmptyCell;
                    break;
                case LocatorKind.Paragraph:
                    spec.Kind = InsertTargetLocatorHelper.LocatorKind.Paragraph;
                    break;
                default:
                    spec.Kind = InsertTargetLocatorHelper.LocatorKind.Sentence;
                    break;
            }

            return spec;
        }

        private static ReplaceExecutionResult ExecuteReplaceWithPayload(
            Word.Document doc,
            ParagraphAction action,
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            ActionContentHelper.Payload payload = action?.ContentPayload;
            if (payload == null || !payload.Ok)
            {
                return new ReplaceExecutionResult
                {
                    Success = false,
                    ErrorMessage = payload?.Error ?? "replace.new 无效"
                };
            }

            if (payload.IsRuns)
            {
                return ExecuteReplaceRuns(
                    doc,
                    action.SentenceNames,
                    action.OriginalScopeFields?.TableId,
                    payload,
                    tableScope,
                    disallowBackendApi);
            }

            if (!ActionFormatHelper.TryBuildSnapshot(
                doc, payload.FormatSpec, tableScope, out var snapshot, out var formatMode, out string fmtErr))
            {
                return new ReplaceExecutionResult { Success = false, ErrorMessage = fmtErr };
            }

            return ExecuteReplaceBySentenceNames(
                doc,
                action.SentenceNames,
                payload.Text,
                action.OriginalScopeFields?.TableId,
                snapshot,
                formatMode,
                disallowBackendApi);
        }

        private static InsertExecutionResult ExecuteInsertWithPayload(
            Word.Document doc,
            ParagraphAction action,
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            ActionContentHelper.Payload payload = action?.ContentPayload;
            if (payload == null || !payload.Ok)
            {
                return new InsertExecutionResult
                {
                    Success = false,
                    ErrorMessage = payload?.Error ?? "insert 文本无效",
                    NewSentences = new List<string>()
                };
            }

            if (payload.IsRuns)
            {
                return ExecuteInsertRuns(
                    doc, action.TargetLocator, payload, tableScope, disallowBackendApi);
            }

            if (!ActionFormatHelper.TryBuildSnapshot(
                doc, payload.FormatSpec, tableScope, out var snapshot, out var formatMode, out string fmtErr))
            {
                return new InsertExecutionResult
                {
                    Success = false,
                    ErrorMessage = fmtErr,
                    NewSentences = new List<string>()
                };
            }

            return ExecuteInsert(
                doc,
                action.TargetLocator,
                payload.Text,
                snapshot,
                formatMode,
                disallowBackendApi);
        }

        private static ReplaceExecutionResult ExecuteReplaceRuns(
            Word.Document doc,
            string sentenceNames,
            string tableId,
            ActionContentHelper.Payload payload,
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            var result = new ReplaceExecutionResult { Success = false, NewSentences = new List<string>() };
            var names = ParseCodeList(sentenceNames);
            if (names.Count == 1 && EmptyCellCode.TryParse(names[0], out EmptyCellCode.Parsed emptyCell))
            {
                return ExecuteEmptyCellReplaceRuns(doc, emptyCell, payload, tableScope, disallowBackendApi);
            }

            Dictionary<string, object> firstSnap;
            string firstMode;
            if (!TrySnapshotForRun(doc, payload.Runs[0], tableScope, out firstSnap, out firstMode, out string err))
            {
                result.ErrorMessage = err;
                return result;
            }

            string joined = payload.JoinedText;
            var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                doc,
                new InsertTargetLocatorHelper.LocatorSpec
                {
                    Kind = SpanLocatorKind(sentenceNames),
                    Codes = sentenceNames,
                    InTableId = tableId
                },
                InsertTargetLocatorHelper.LocateRangeKind.ContentSpan,
                disallowBackendApi,
                refreshDocument: false);
            if (!locate.Success || locate.Range == null)
            {
                result.ErrorMessage = locate.Error ?? "未找到任何句子 Range";
                return result;
            }

            Word.Range targetRange = locate.Range;
            string originalRangeText = targetRange.Text ?? "";
            string aligned = TextUtils.EnsureReplaceContentTrailingParagraphMarks(joined, originalRangeText);
            if (aligned != joined && payload.Runs.Count > 0)
            {
                ActionContentHelper.Run last = payload.Runs[payload.Runs.Count - 1];
                last.Text += aligned.Substring(joined.Length);
            }

            string trailingSuffix = TextUtils.GetTrailingParagraphSuffix(originalRangeText);
            int trailingLen = trailingSuffix.Length;
            if (trailingLen > 0)
            {
                int rangeSpan = targetRange.End - targetRange.Start;
                if (rangeSpan > trailingLen)
                {
                    targetRange = doc.Range(targetRange.Start, targetRange.End - trailingLen);
                }

                payload.Runs[payload.Runs.Count - 1].Text =
                    TextUtils.TrimTrailingParagraphMarks(payload.Runs[payload.Runs.Count - 1].Text);
            }

            result.OriginalCodes = names;
            string writeErr;
            if (!TryWriteRuns(
                doc,
                targetRange,
                payload.Runs,
                tableScope,
                replaceFirst: true,
                WrittenFormatGlueHelper.IsPrefixSafe(sentenceNames, insertAfter: false),
                out writeErr,
                out List<string> newSentences))
            {
                result.ErrorMessage = writeErr;
                return result;
            }

            result.NewSentences = newSentences;
            result.FormatMode = firstMode;
            result.Success = true;
            return result;
        }

        private static InsertExecutionResult ExecuteInsertRuns(
            Word.Document doc,
            LocatorSpec target,
            ActionContentHelper.Payload payload,
            TableScopeIndex tableScope,
            bool disallowBackendApi)
        {
            var result = new InsertExecutionResult { Success = false, NewSentences = new List<string>() };
            if (target != null && target.Kind == LocatorKind.EmptyCell)
            {
                EmptyCellCode.Parsed parsed = new EmptyCellCode.Parsed
                {
                    TableId = target.TableId,
                    Row = target.Row,
                    Col = target.Col,
                    Code = target.Codes
                };
                var written = ExecuteEmptyCellReplaceRuns(doc, parsed, payload, tableScope, disallowBackendApi);
                result.Success = written.Success;
                result.ErrorMessage = written.ErrorMessage;
                result.FormatMode = written.FormatMode;
                result.FormatApplyError = written.FormatApplyError;
                result.NewSentences = written.NewSentences ?? new List<string>();
                return result;
            }

            var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                doc,
                ToFacadeSpec(target),
                InsertTargetLocatorHelper.LocateRangeKind.InsertAfter,
                disallowBackendApi,
                refreshDocument: false);
            if (!locate.Success || locate.Range == null)
            {
                result.ErrorMessage = locate.Error ?? "无法确定插入位置";
                return result;
            }

            if (!TryWriteRuns(
                doc,
                locate.Range,
                payload.Runs,
                tableScope,
                replaceFirst: false,
                PrefixSafeForInsertTarget(target),
                out string writeErr,
                out List<string> newSentences))
            {
                result.ErrorMessage = writeErr;
                return result;
            }

            result.NewSentences = newSentences;
            result.FormatMode = "explicit";
            result.Success = true;
            return result;
        }

        private static ReplaceExecutionResult ExecuteEmptyCellReplaceRuns(
            Word.Document doc,
            EmptyCellCode.Parsed emptyCell,
            ActionContentHelper.Payload payload,
            TableScopeIndex tableScope,
            bool disallowBackendApi,
            Word.Cell resolvedCell = null)
        {
            var result = new ReplaceExecutionResult { Success = false, NewSentences = new List<string>() };
            result.OriginalCodes = new List<string> { emptyCell.Code };

            Word.Cell cell = resolvedCell;
            if (cell == null)
            {
                DocumentActionPathLog.Record(
                    DocumentActionPathLog.LogTag + " locate via=single n=1");
                var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                    doc,
                    new InsertTargetLocatorHelper.LocatorSpec
                    {
                        Kind = InsertTargetLocatorHelper.LocatorKind.EmptyCell,
                        Codes = emptyCell.Code,
                        TableId = emptyCell.TableId,
                        Row = emptyCell.Row,
                        Col = emptyCell.Col
                    },
                    InsertTargetLocatorHelper.LocateRangeKind.WriteCell,
                    disallowBackendApi,
                    refreshDocument: false);
                if (!locate.Success)
                {
                    result.ErrorMessage = locate.Error ?? "无法定位空格";
                    return result;
                }

                Word.Table table = TableResolveHelper.TryResolveTableById(doc, emptyCell.TableId, out string tableError);
                if (table == null)
                {
                    result.ErrorMessage = tableError ?? "找不到表格";
                    return result;
                }

                cell = TableRowValuesHelper.TryResolveWritableCell(table, emptyCell.Row, emptyCell.Col);
                if (cell == null)
                {
                    result.ErrorMessage = TableRowValuesHelper.BuildUnresolvableCellError(emptyCell.Row, emptyCell.Col);
                    return result;
                }
            }

            ActionContentHelper.Run first = payload.Runs[0];
            if (!TrySnapshotForRun(doc, first, tableScope, out var firstSnap, out var firstMode, out string fmtErr))
            {
                result.ErrorMessage = fmtErr;
                return result;
            }

            CellWriteResult write = TableCellContentWriter.WriteCell(cell, first.Text ?? "");
            if (!write.Success)
            {
                result.ErrorMessage = write.Error ?? "写入单元格失败";
                return result;
            }

            try
            {
                FormatInheritHelper.ApplySnapshot(cell.Range, firstSnap, charFormatOnly: true);
            }
            catch (Exception fmtEx)
            {
                result.FormatApplyError = fmtEx.Message;
                result.ErrorMessage = $"文字已写入，但字符格式套用失败: {fmtEx.Message}";
                return result;
            }

            var rest = new List<ActionContentHelper.Run>();
            for (int i = 1; i < payload.Runs.Count; i++)
            {
                rest.Add(payload.Runs[i]);
            }

            Word.Range after = cell.Range;
            after.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            if (after.Start > cell.Range.Start)
            {
                after = doc.Range(Math.Max(cell.Range.Start, after.Start - 1), Math.Max(cell.Range.Start, after.Start - 1));
            }

            var sentences = new List<string>(SentenceSplitter.SplitForActionContent(first.Text ?? ""));
            if (rest.Count > 0)
            {
                if (!TryWriteRuns(
                    doc,
                    after,
                    rest,
                    tableScope,
                    replaceFirst: false,
                    WrittenFormatGlueHelper.PreviousRunEndsWithBreak(first.Text),
                    out string writeErr,
                    out List<string> more))
                {
                    result.ErrorMessage = writeErr;
                    return result;
                }

                sentences.AddRange(more);
            }

            MoveSelectionAfterReplace(doc, cell.Range);
            result.NewSentences = sentences;
            result.FormatMode = firstMode;
            result.Success = true;
            return result;
        }

        private static bool TrySnapshotForRun(
            Word.Document doc,
            ActionContentHelper.Run run,
            TableScopeIndex tableScope,
            out Dictionary<string, object> snapshot,
            out string formatMode,
            out string error)
        {
            return ActionFormatHelper.TryBuildSnapshot(
                doc, run.FormatSpec, tableScope, out snapshot, out formatMode, out error);
        }

        private static bool TryWriteRuns(
            Word.Document doc,
            Word.Range startRange,
            List<ActionContentHelper.Run> runs,
            TableScopeIndex tableScope,
            bool replaceFirst,
            bool firstPrefixSafe,
            out string error,
            out List<string> newSentences)
        {
            error = null;
            newSentences = new List<string>();
            Word.Range cursor = startRange;
            for (int i = 0; i < runs.Count; i++)
            {
                ActionContentHelper.Run run = runs[i];
                if (!TrySnapshotForRun(doc, run, tableScope, out var snap, out _, out error))
                {
                    return false;
                }

                Word.Range writeAt = cursor;
                if (i > 0 || !replaceFirst)
                {
                    writeAt = doc.Range(cursor.End, cursor.End);
                }

                Word.Range written = WordBreakWriteHelper.WriteTextWithBreaks(doc, writeAt, run.Text)
                    ?? writeAt;
                try
                {
                    bool prefixSafe = i == 0
                        ? firstPrefixSafe
                        : WrittenFormatGlueHelper.PreviousRunEndsWithBreak(runs[i - 1].Text);
                    FormatInheritHelper.ApplyWrittenCharFormat(
                        doc, written, run.Text ?? "", snap, prefixSafe, FindSentenceInRange);
                }
                catch (Exception fmtEx)
                {
                    error = "文字已写入，但字符格式套用失败: " + fmtEx.Message;
                    return false;
                }

                newSentences.AddRange(SentenceSplitter.SplitForActionContent(run.Text ?? ""));
                cursor = written;
            }

            MoveSelectionAfterInsert(doc, cursor);
            return true;
        }

        private static ReplaceExecutionResult ExecuteEmptyCellReplace(
            Word.Document doc,
            EmptyCellCode.Parsed emptyCell,
            string newContent,
            Dictionary<string, object> formatSnapshot,
            string formatMode,
            bool disallowBackendApi,
            Word.Cell resolvedCell = null)
        {
            var result = new ReplaceExecutionResult { Success = false, NewSentences = new List<string>() };
            if (emptyCell == null)
            {
                result.ErrorMessage = "空格码无效";
                return result;
            }

            result.OriginalCodes = new List<string> { emptyCell.Code };
            if (formatSnapshot == null || formatSnapshot.Count == 0 || string.IsNullOrEmpty(formatMode))
            {
                result.ErrorMessage = ActionFormatHelper.MissingFormatError;
                return result;
            }

            Word.Cell cell = resolvedCell;
            if (cell == null)
            {
                DocumentActionPathLog.Record(
                    DocumentActionPathLog.LogTag + " locate via=single n=1");
                var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                    doc,
                    new InsertTargetLocatorHelper.LocatorSpec
                    {
                        Kind = InsertTargetLocatorHelper.LocatorKind.EmptyCell,
                        Codes = emptyCell.Code,
                        TableId = emptyCell.TableId,
                        Row = emptyCell.Row,
                        Col = emptyCell.Col
                    },
                    InsertTargetLocatorHelper.LocateRangeKind.WriteCell,
                    disallowBackendApi,
                    refreshDocument: false);
                if (!locate.Success)
                {
                    result.ErrorMessage = locate.Error ?? "无法定位空格";
                    return result;
                }

                Word.Table table = TableResolveHelper.TryResolveTableById(doc, emptyCell.TableId, out string tableError);
                if (table == null)
                {
                    result.ErrorMessage = tableError ?? "找不到表格";
                    return result;
                }

                cell = TableRowValuesHelper.TryResolveWritableCell(table, emptyCell.Row, emptyCell.Col);
                if (cell == null)
                {
                    result.ErrorMessage = TableRowValuesHelper.BuildUnresolvableCellError(emptyCell.Row, emptyCell.Col);
                    return result;
                }
            }

            CellWriteResult write = TableCellContentWriter.WriteCell(cell, newContent ?? "");
            if (!write.Success)
            {
                result.ErrorMessage = write.Error ?? "写入单元格失败";
                return result;
            }

            try
            {
                FormatInheritHelper.ApplySnapshot(cell.Range, formatSnapshot, charFormatOnly: true);
                result.FormatMode = formatMode;
            }
            catch (Exception fmtEx)
            {
                result.FormatApplyError = fmtEx.Message;
                result.ErrorMessage = $"文字已写入，但字符格式套用失败: {fmtEx.Message}";
                return result;
            }

            MoveSelectionAfterReplace(doc, cell.Range);
            result.NewSentences = SentenceSplitter.SplitForActionContent(newContent ?? "");
            result.Success = true;
            return result;
        }

        private static DeleteExecutionResult ExecuteEmptyCellClear(
            Word.Document doc,
            EmptyCellCode.Parsed emptyCell,
            bool disallowBackendApi,
            Word.Cell resolvedCell = null)
        {
            var result = new DeleteExecutionResult { Success = false };
            if (emptyCell == null)
            {
                result.ErrorMessage = "空格码无效";
                return result;
            }

            result.AffectedCodes = new List<string> { emptyCell.Code };
            Word.Cell cell = resolvedCell;
            if (cell == null)
            {
                DocumentActionPathLog.Record(
                    DocumentActionPathLog.LogTag + " locate via=single n=1");
                var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                    doc,
                    new InsertTargetLocatorHelper.LocatorSpec
                    {
                        Kind = InsertTargetLocatorHelper.LocatorKind.EmptyCell,
                        Codes = emptyCell.Code,
                        TableId = emptyCell.TableId,
                        Row = emptyCell.Row,
                        Col = emptyCell.Col
                    },
                    InsertTargetLocatorHelper.LocateRangeKind.WriteCell,
                    disallowBackendApi,
                    refreshDocument: false);
                if (!locate.Success)
                {
                    result.ErrorMessage = locate.Error ?? "无法定位空格";
                    return result;
                }

                Word.Table table = TableResolveHelper.TryResolveTableById(doc, emptyCell.TableId, out string tableError);
                if (table == null)
                {
                    result.ErrorMessage = tableError ?? "找不到表格";
                    return result;
                }

                cell = TableRowValuesHelper.TryResolveWritableCell(table, emptyCell.Row, emptyCell.Col);
                if (cell == null)
                {
                    result.ErrorMessage = TableRowValuesHelper.BuildUnresolvableCellError(emptyCell.Row, emptyCell.Col);
                    return result;
                }
            }

            CellWriteResult write = TableCellContentWriter.WriteCell(cell, "");
            if (!write.Success)
            {
                result.ErrorMessage = write.Error ?? "清空单元格失败";
                return result;
            }

            MoveSelectionAfterReplace(doc, cell.Range);
            result.Success = true;
            return result;
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
            string tableId,
            Dictionary<string, object> formatSnapshot,
            string formatMode,
            bool disallowBackendApi)
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

                var names = ParseCodeList(sentenceNames);
                if (names.Count == 0)
                {
                    result.ErrorMessage = "original.codes 为空";
                    return result;
                }

                if (names.Count == 1 && EmptyCellCode.TryParse(names[0], out EmptyCellCode.Parsed emptyCell))
                {
                    return ExecuteEmptyCellReplace(
                        doc, emptyCell, newContent, formatSnapshot, formatMode, disallowBackendApi);
                }

                if (!LooksLikeParagraphCodes(sentenceNames)
                    && names.Any(n => n.StartsWith("T_", StringComparison.Ordinal)
                    || n.StartsWith("C_", StringComparison.Ordinal)))
                {
                    result.ErrorMessage = "replace 仅支持 S_ 编码；T_/C_ 仅用于 insert 锚点";
                    return result;
                }

                result.OriginalCodes = names;
                FormatInheritHelper.DbgLog($"ExecuteReplace S_ count={names.Count} codes=[{string.Join(",", names)}]");

                var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                    doc,
                    new InsertTargetLocatorHelper.LocatorSpec
                    {
                        Kind = SpanLocatorKind(sentenceNames),
                        Codes = sentenceNames,
                        InTableId = tableId
                    },
                    InsertTargetLocatorHelper.LocateRangeKind.ContentSpan,
                    disallowBackendApi,
                    refreshDocument: false);
                if (!locate.Success || locate.Range == null)
                {
                    result.ErrorMessage = locate.Error ?? "未找到任何句子 Range";
                    System.Diagnostics.Debug.WriteLine($"[替换操作] {result.ErrorMessage}");
                    return result;
                }

                Word.Range targetRange = locate.Range;

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
                    WordBreakWriteHelper.ContainsLayoutBreaks(replacementText)
                        ? "ExecuteReplace 即将分段写入（\\f / ^b → InsertBreak）"
                        : WordNoteMarkHelper.ContainsNoteMarks(replacementText)
                            ? "ExecuteReplace 即将写入（含 ^f/^e → 脚注/尾注引用）"
                            : "ExecuteReplace 即将 targetRange.Text=...（会清除该区间直接格式）");
                Word.Range writtenRange = WordBreakWriteHelper.WriteTextWithBreaks(
                    doc, targetRange, replacementText);
                targetRange = writtenRange ?? targetRange;
                FormatInheritHelper.DbgLogRangeFonts("ExecuteReplace targetRange AFTER Text=", targetRange);

                Word.Range newRange = doc.Range(targetRange.Start, targetRange.End);

                try
                {
                    FormatInheritHelper.DbgLog($"ExecuteReplace ApplySnapshot format_mode={formatMode}");
                    FormatInheritHelper.ApplyWrittenCharFormat(
                        doc,
                        newRange,
                        replacementText,
                        formatSnapshot,
                        WrittenFormatGlueHelper.IsPrefixSafe(sentenceNames, insertAfter: false),
                        FindSentenceInRange);
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
            string tableId,
            bool disallowBackendApi)
        {
            var result = new DeleteExecutionResult { Success = false };

            try
            {
                if (string.IsNullOrEmpty(sentenceNames))
                {
                    result.ErrorMessage = "缺少 deletion.codes";
                    return result;
                }

                var names = ParseCodeList(sentenceNames);
                if (names.Count == 0)
                {
                    result.ErrorMessage = "deletion.codes 为空";
                    return result;
                }

                if (names.Count == 1 && EmptyCellCode.TryParse(names[0], out EmptyCellCode.Parsed emptyCell))
                {
                    return ExecuteEmptyCellClear(doc, emptyCell, disallowBackendApi);
                }

                if (!LooksLikeParagraphCodes(sentenceNames)
                    && names.Any(n => n.StartsWith("T_", StringComparison.Ordinal)
                    || n.StartsWith("C_", StringComparison.Ordinal)))
                {
                    result.ErrorMessage = "delete 支持 S_ 或一条几何码 T_…_r_c；光杆 T_/C_ 请用专用工具";
                    return result;
                }

                result.AffectedCodes = names;

                var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                    doc,
                    new InsertTargetLocatorHelper.LocatorSpec
                    {
                        Kind = SpanLocatorKind(sentenceNames),
                        Codes = sentenceNames,
                        InTableId = tableId
                    },
                    InsertTargetLocatorHelper.LocateRangeKind.ContentSpan,
                    disallowBackendApi,
                    refreshDocument: false);
                if (!locate.Success || locate.Range == null)
                {
                    result.ErrorMessage = locate.Error ?? "未找到任何句子 Range";
                    System.Diagnostics.Debug.WriteLine($"[删除操作] {result.ErrorMessage}");
                    return result;
                }

                Word.Range targetRange = locate.Range;

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

        /// <summary>
        /// 按 target 定位后插入（仅后插）：门面 InsertAfter。
        /// </summary>
        private static InsertExecutionResult ExecuteInsert(
            Word.Document doc,
            LocatorSpec target,
            string insertContent,
            Dictionary<string, object> formatSnapshot,
            string formatMode,
            bool disallowBackendApi)
        {
            var result = new InsertExecutionResult { Success = false, NewSentences = new List<string>() };

            try
            {
                if (string.IsNullOrEmpty(insertContent))
                {
                    result.ErrorMessage = "缺少 insert 内容";
                    return result;
                }

                if (target != null && target.Kind == LocatorKind.EmptyCell)
                {
                    EmptyCellCode.Parsed parsed = new EmptyCellCode.Parsed
                    {
                        TableId = target.TableId,
                        Row = target.Row,
                        Col = target.Col,
                        Code = target.Codes
                    };
                    var written = ExecuteEmptyCellReplace(
                        doc, parsed, insertContent, formatSnapshot, formatMode, disallowBackendApi);
                    result.Success = written.Success;
                    result.ErrorMessage = written.ErrorMessage;
                    result.FormatMode = written.FormatMode;
                    result.FormatApplyError = written.FormatApplyError;
                    result.NewSentences = written.NewSentences ?? new List<string>();
                    return result;
                }

                var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                    doc,
                    ToFacadeSpec(target),
                    InsertTargetLocatorHelper.LocateRangeKind.InsertAfter,
                    disallowBackendApi,
                    refreshDocument: false);
                if (!locate.Success || locate.Range == null)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "insert-no-range",
                        "kind=" + (target == null ? "cursor" : target.Kind.ToString())
                        + " err=" + (locate.Error ?? ""));
                    result.ErrorMessage = locate.Error ?? "无法确定插入位置";
                    return result;
                }

                Word.Range insertRange = locate.Range;
                System.Diagnostics.Debug.WriteLine(
                    $"[ProcessActions][insert] 门面定位 usedCursor={locate.UsedCursor} kind={target?.Kind}");

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
                insertRange = WordBreakWriteHelper.WriteTextWithBreaks(doc, insertRange, insertText)
                    ?? insertRange;
                MoveSelectionAfterInsert(doc, insertRange);

                try
                {
                    Word.Range insertedBlock = doc.Range(insertRange.Start, insertRange.End);
                    FormatInheritHelper.ApplyWrittenCharFormat(
                        doc,
                        insertedBlock,
                        insertText,
                        formatSnapshot,
                        PrefixSafeForInsertTarget(target),
                        FindSentenceInRange);
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

        private static bool PrefixSafeForInsertTarget(LocatorSpec target)
        {
            if (target == null
                || target.Kind == LocatorKind.DocumentStart
                || target.Kind == LocatorKind.EmptyCell)
            {
                return true;
            }

            if (target.Kind != LocatorKind.Sentence || string.IsNullOrEmpty(target.Codes))
            {
                return true;
            }

            return WrittenFormatGlueHelper.IsPrefixSafe(target.Codes, insertAfter: true);
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
                string convertedText = WordRangeFinder.ConvertNewlinesToWordCodes(sentenceText);
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
            string pattern = @"\[S_[0-9a-zA-Z]{5}-start(?: [^\]]*)?\]|\[S_[0-9a-zA-Z]{5}-end\]";
            string result = Regex.Replace(text, pattern, "");
            return WordAddIn1.DocumentMapping.InTableBoundaryMarker.Strip(result);
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

            /// <summary>replace/insert 的 detail.format；delete 为 null。分片时为 null。</summary>
            public ActionFormatHelper.Spec FormatSpec { get; set; }

            public ActionContentHelper.Payload ContentPayload { get; set; }
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

        private static bool LooksLikeParagraphCodes(string codes)
        {
            List<string> names = ParseCodeList(codes);
            return names.Count > 0 && FormatContextHelper.IsParagraphCode(names[0]);
        }

        private static InsertTargetLocatorHelper.LocatorKind SpanLocatorKind(string codes)
        {
            return LooksLikeParagraphCodes(codes)
                ? InsertTargetLocatorHelper.LocatorKind.Paragraph
                : InsertTargetLocatorHelper.LocatorKind.Sentence;
        }

        private static async Task RemapAndAssignNewCodes(
            Word.Document doc,
            ActionResult result,
            bool disallowBackendApi,
            string pdMode = "single")
        {
            if (result == null || !result.success)
            {
                return;
            }

            try
            {
                DocumentActionPathLog.Record(
                    DocumentActionPathLog.LogTag + " post_pd mode=" + (pdMode ?? "single"));
                using (EasyWriteDiagnostics.Time(
                    "process_actions.post_pd[" + result.action_index + "]"))
                {
                    string postError = await UpdateSentenceNameMapping(
                        doc,
                        "process_actions_step4",
                        disallowBackendApi);
                    if (!string.IsNullOrEmpty(postError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            "[process_actions] 写后 PD warning: " + postError);
                    }
                }

                // 对外已不回 new_codes；写后 PD 只刷新映射。邻居夹码先停。
                // ApplySnapshotDiffToResults(new List<ActionResult> { result });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[process_actions] 写后 PD 异常: " + ex.Message);
            }
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

        /// <summary>
        /// 写前快照：step4 之前最近一条即可。
        /// 现网 precheck 用 source=locate，缓存命中且快照未变时不入栈；
        /// 栈上常见的是 get_document_content。只认 step1/locate 会得到 snapshot_history_missing。
        /// </summary>
        private static SnapshotRecord FindLatestStep1Before(int postIndex)
        {
            for (int i = DocumentState.SnapshotHistory.Count - 1; i >= 0; i--)
            {
                SnapshotRecord record = DocumentState.SnapshotHistory[i];
                if (record.Index < postIndex)
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
            // 写入已成功时保持「替换/插入成功」，不要催读文，避免模型当成失败再改。
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
            var affectedUnion = new List<string>();
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
            }

            return new Dictionary<string, object>
            {
                ["affected_codes"] = affectedUnion
            };
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

            if (rep.FormatMode == "explicit")
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
                if (target.Kind == LocatorKind.Sentence || target.Kind == LocatorKind.Paragraph)
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
                else if (target.Kind == LocatorKind.EmptyCell)
                {
                    anchorCodes = target.Codes;
                }
            }

            ar.affected_codes = ParseCodeList(anchorCodes);
            ar.PendingNewSentences = ins.NewSentences ?? new List<string>();
            ar.InsertAnchorCodes = ParseCodeList(anchorCodes);
            ar.success = ins.Success;
            ar.format_mode = ins.FormatMode;
            ar.format_apply_error = ins.FormatApplyError;
            if (ins.FormatMode == "explicit")
            {
                ar.format_applied = "char_format";
                ar.format_source = "detail.format.char_format";
            }

            if (!ins.Success)
            {
                ar.message = ins.ErrorMessage ?? "插入失败";
            }
            else if (target != null && target.Kind == LocatorKind.DocumentStart)
            {
                ar.message = "已插在正文最前";
            }
            else
            {
                ar.message = "插入成功";
            }
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

