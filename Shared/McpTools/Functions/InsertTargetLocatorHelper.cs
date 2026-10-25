using System;
using System.Collections.Generic;
using System.Linq;
using WordAddIn1.DocumentHost;
using WordAddIn1.DocumentMapping;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 统一定位：解析 → 按需 ProcessDocument → 歧义（含消歧建议）→ Range。
    /// </summary>
    public static class InsertTargetLocatorHelper
    {
        public enum LocatorKind
        {
            Sentence,
            Table,
            Chart,
            Image,
            DocumentStart,
            EmptyCell,
            Paragraph
        }

        public enum LocateRangeKind
        {
            InsertAfter = 0,
            ContentSpan = 1,
            WriteCell = 2
        }

        public enum LocatePrecheckStream
        {
            SentenceActions = 0,
            SentenceFormatApply = 1,
            SentenceFormatContext = 2,
            ParagraphFormatApply = 3,
            ParagraphFormatContext = 4
        }

        public const string PositionStart = "start";
        public const string FiveWayLocatorError =
            "定位对象须为 codes / paragraph_codes / table / chart / image / position 六选一";
        public const string SentenceOnlyLocatorError = "仅支持 codes / paragraph_codes";
        public const string InTableWithCodesOrParagraphError = "in_table 仅可与 codes / paragraph_codes 同用";
        public const string CodesRejectParagraphError = "codes 仅支持 S_ 或真空几何码；段落请用 paragraph_codes";
        public const string ParagraphCodesOnlyPError = "paragraph_codes 仅支持 P_ 段落编码";
        public const string FromFileRetiredFlatLocatorError =
            "请改用嵌套 target / original（与 process_actions 同构）；已废止 target_codes / original_codes / in_table 等扁平定位键";

        public static readonly string[] FromFileRetiredInsertKeys =
        {
            "target_codes", "in_table", "target_table", "target_chart", "target_image", "target_position"
        };

        public static readonly string[] FromFileRetiredReplaceKeys =
        {
            "original_codes", "in_table"
        };

        public static string RetiredFlatLocatorError(Dictionary<string, object> args, IReadOnlyList<string> keys)
        {
            if (args == null || keys == null)
            {
                return null;
            }

            foreach (string key in keys)
            {
                if (args.ContainsKey(key) && args[key] != null
                    && !string.IsNullOrWhiteSpace(args[key].ToString()))
                {
                    return FromFileRetiredFlatLocatorError;
                }
            }

            return null;
        }
        public const string PositionStartOnlyError = "position 仅支持 start";

        public sealed class LocatorSpec
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

        public sealed class LocatorParseResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public LocatorSpec Spec { get; set; }
        }

        public sealed class ResolveRangeResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public Word.Range Range { get; set; }
            public bool UsedCursor { get; set; }
            public LocatorSpec Spec { get; set; }
            public AmbiguityCheckResult Ambiguity { get; set; }
        }

        public sealed class PrecheckResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public AmbiguityCheckResult Ambiguity { get; set; }
            public TableScopeIndex TableScope { get; set; }
        }

        /// <summary>
        /// 解析定位对象：codes / paragraph_codes / table / chart / image / position 六选一。
        /// in_table 仅可与 codes 或段落字段同用（schema 不暴露，内部仍认）。
        /// </summary>
        public static LocatorParseResult ParseLocatorObject(
            Dictionary<string, object> dict,
            bool sentenceOnly,
            string fieldLabel)
        {
            if (dict == null)
            {
                return new LocatorParseResult { Success = false, Error = $"{fieldLabel} 定位对象为空" };
            }

            bool hasCodes = TryGetLocatorField(dict, "codes", out _);
            bool hasParaNested = TryGetLocatorField(dict, "paragraph_codes", out _);
            bool hasParaFlat = TryGetLocatorField(dict, "target_paragraph_codes", out _);
            if (hasParaNested && hasParaFlat)
            {
                return FailParse($"{fieldLabel} paragraph_codes 与 target_paragraph_codes 不能同时出现");
            }

            bool hasPara = hasParaNested || hasParaFlat;
            bool hasTable = TryGetLocatorField(dict, "table", out _);
            bool hasChart = TryGetLocatorField(dict, "chart", out _);
            bool hasImage = TryGetLocatorField(dict, "image", out _);
            bool hasPosition = TryGetLocatorField(dict, "position", out _);
            bool hasInTable = TryGetLocatorField(dict, "in_table", out _);

            int modeCount = (hasCodes ? 1 : 0) + (hasPara ? 1 : 0) + (hasTable ? 1 : 0)
                + (hasChart ? 1 : 0) + (hasImage ? 1 : 0) + (hasPosition ? 1 : 0);
            if (modeCount != 1)
            {
                return FailParse($"{fieldLabel} {FiveWayLocatorError}");
            }

            if (sentenceOnly && (hasTable || hasChart || hasImage || hasPosition))
            {
                return FailParse($"{fieldLabel} {SentenceOnlyLocatorError}");
            }

            if (hasPosition)
            {
                if (hasInTable)
                {
                    return FailParse($"{fieldLabel}.{InTableWithCodesOrParagraphError}");
                }

                if (!IsStartPosition(dict["position"].ToString()))
                {
                    return FailParse($"{fieldLabel}.{PositionStartOnlyError}");
                }

                return new LocatorParseResult
                {
                    Success = true,
                    Spec = new LocatorSpec { Kind = LocatorKind.DocumentStart }
                };
            }

            if (hasCodes)
            {
                return ParseCodesMode(dict, fieldLabel, hasInTable);
            }

            if (hasPara)
            {
                return ParseParagraphMode(dict, fieldLabel, hasParaNested, hasInTable);
            }

            if (hasInTable)
            {
                return FailParse($"{fieldLabel}.{InTableWithCodesOrParagraphError}");
            }

            if (hasTable)
            {
                string tableId = dict["table"].ToString().Trim();
                if (EmptyCellCode.IsEmptyCellCode(tableId))
                {
                    return new LocatorParseResult
                    {
                        Success = false,
                        Error = $"{fieldLabel}.{EmptyCellCode.TableIdNotCellError}"
                    };
                }

                if (!tableId.StartsWith("T_", StringComparison.Ordinal))
                {
                    return new LocatorParseResult
                    {
                        Success = false,
                        Error = $"{fieldLabel}.table 须为单个 T_ 表格编号"
                    };
                }

                return new LocatorParseResult
                {
                    Success = true,
                    Spec = new LocatorSpec { Kind = LocatorKind.Table, TableId = tableId }
                };
            }

            if (hasChart)
            {
                string chartId = dict["chart"].ToString().Trim();
                if (!chartId.StartsWith("C_", StringComparison.Ordinal))
                {
                    return new LocatorParseResult
                    {
                        Success = false,
                        Error = $"{fieldLabel}.chart 须为单个 C_ 图表编号"
                    };
                }

                return new LocatorParseResult
                {
                    Success = true,
                    Spec = new LocatorSpec { Kind = LocatorKind.Chart, ChartId = chartId }
                };
            }

            string imageId = dict["image"].ToString().Trim();
            if (!imageId.StartsWith("I_", StringComparison.Ordinal))
            {
                return new LocatorParseResult
                {
                    Success = false,
                    Error = $"{fieldLabel}.image 须为单个 I_ 图片编号"
                };
            }

            return new LocatorParseResult
            {
                Success = true,
                Spec = new LocatorSpec { Kind = LocatorKind.Image, ImageId = imageId }
            };
        }

        private static LocatorParseResult FailParse(string error)
        {
            return new LocatorParseResult { Success = false, Error = error };
        }

        private static bool TryGetLocatorField(Dictionary<string, object> dict, string key, out string value)
        {
            value = null;
            if (dict == null || !dict.ContainsKey(key) || dict[key] == null)
            {
                return false;
            }

            value = dict[key].ToString().Trim();
            return !string.IsNullOrEmpty(value);
        }

        private static LocatorParseResult ParseInTableId(
            Dictionary<string, object> dict,
            bool hasInTable,
            string fieldLabel,
            out string inTable)
        {
            inTable = null;
            if (!hasInTable)
            {
                return new LocatorParseResult { Success = true };
            }

            inTable = dict["in_table"].ToString().Trim();
            if (!inTable.StartsWith("T_", StringComparison.Ordinal))
            {
                return FailParse($"{fieldLabel}.in_table 须为 T_ 表格编号");
            }

            if (EmptyCellCode.IsEmptyCellCode(inTable))
            {
                return FailParse($"{fieldLabel}.in_table 须为光杆表号，不能用空格码");
            }

            return new LocatorParseResult { Success = true };
        }

        private static LocatorParseResult ParseCodesMode(
            Dictionary<string, object> dict,
            string fieldLabel,
            bool hasInTable)
        {
            string codes = dict["codes"].ToString().Trim();
            var parts = codes.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
            if (parts.Count == 0)
            {
                return FailParse($"{fieldLabel}.codes 为空");
            }

            bool anyEmpty = parts.Any(EmptyCellCode.IsEmptyCellCode);
            if (anyEmpty)
            {
                EmptyCellCode.Parsed parsedCell;
                if (parts.Count != 1 || !EmptyCellCode.TryParse(parts[0], out parsedCell))
                {
                    return FailParse($"{fieldLabel}.codes {EmptyCellCode.MixError}");
                }

                string inTableEmpty = hasInTable ? dict["in_table"].ToString().Trim() : null;
                if (!string.IsNullOrEmpty(inTableEmpty)
                    && !string.Equals(inTableEmpty, parsedCell.TableId, StringComparison.Ordinal))
                {
                    return FailParse($"{fieldLabel}.in_table 须等于空格码内表号 {parsedCell.TableId}");
                }

                return new LocatorParseResult
                {
                    Success = true,
                    Spec = new LocatorSpec
                    {
                        Kind = LocatorKind.EmptyCell,
                        Codes = parsedCell.Code,
                        TableId = parsedCell.TableId,
                        InTableId = string.IsNullOrEmpty(inTableEmpty) ? parsedCell.TableId : inTableEmpty,
                        Row = parsedCell.Row,
                        Col = parsedCell.Col
                    }
                };
            }

            foreach (string part in parts)
            {
                if (part.StartsWith("P_", StringComparison.Ordinal))
                {
                    return FailParse($"{fieldLabel}.{CodesRejectParagraphError}");
                }

                if (part.StartsWith("T_", StringComparison.Ordinal)
                    || part.StartsWith("C_", StringComparison.Ordinal)
                    || part.StartsWith("I_", StringComparison.Ordinal))
                {
                    return FailParse($"{fieldLabel}.codes 仅支持 S_；表/图/图片请用 table / chart / image 字段");
                }

                if (!part.StartsWith("S_", StringComparison.Ordinal))
                {
                    return FailParse($"{fieldLabel}.codes 仅支持 S_ 句子编码");
                }
            }

            LocatorParseResult inTableParsed = ParseInTableId(dict, hasInTable, fieldLabel, out string inTable);
            if (!inTableParsed.Success)
            {
                return inTableParsed;
            }

            return new LocatorParseResult
            {
                Success = true,
                Spec = new LocatorSpec
                {
                    Kind = LocatorKind.Sentence,
                    Codes = string.Join(",", parts),
                    InTableId = inTable
                }
            };
        }

        private static LocatorParseResult ParseParagraphMode(
            Dictionary<string, object> dict,
            string fieldLabel,
            bool hasParaNested,
            bool hasInTable)
        {
            string key = hasParaNested ? "paragraph_codes" : "target_paragraph_codes";
            string csv = dict[key].ToString().Trim();
            var parts = csv.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
            if (parts.Count == 0)
            {
                return FailParse($"{fieldLabel}.{key} 为空");
            }

            foreach (string part in parts)
            {
                if (!FormatContextHelper.IsParagraphCode(part))
                {
                    return FailParse($"{fieldLabel}.{ParagraphCodesOnlyPError}");
                }
            }

            LocatorParseResult inTableParsed = ParseInTableId(dict, hasInTable, fieldLabel, out string inTable);
            if (!inTableParsed.Success)
            {
                return inTableParsed;
            }

            return new LocatorParseResult
            {
                Success = true,
                Spec = new LocatorSpec
                {
                    Kind = LocatorKind.Paragraph,
                    Codes = string.Join(",", parts),
                    InTableId = inTable
                }
            };
        }

        /// <summary>
        /// 解析 args["target"]（可省略）。省略则光标处。
        /// </summary>
        public static ResolveRangeResult TryResolveInsertRange(
            Word.Document doc,
            Dictionary<string, object> args,
            LocateRangeKind rangeKind = LocateRangeKind.InsertAfter,
            bool disallowBackendApi = false)
        {
            if (doc == null)
            {
                return FailResolve("locate_failed: 文档不可用");
            }

            Dictionary<string, object> targetDict = TryGetTargetDict(args);
            if (targetDict == null)
            {
                return new ResolveRangeResult
                {
                    Success = true,
                    UsedCursor = true,
                    Range = doc.Application.Selection.Range,
                    Spec = null
                };
            }

            LocatorParseResult parsed = ParseLocatorObject(targetDict, sentenceOnly: false, fieldLabel: "target");
            if (!parsed.Success)
            {
                return FailResolve($"locate_failed: {parsed.Error}");
            }

            return TryResolveInsertRange(doc, parsed.Spec, rangeKind, disallowBackendApi, refreshDocument: true);
        }

        public static ResolveRangeResult TryResolveInsertRange(
            Word.Document doc,
            LocatorSpec target,
            LocateRangeKind rangeKind = LocateRangeKind.InsertAfter,
            bool disallowBackendApi = false,
            bool refreshDocument = true)
        {
            if (doc == null)
            {
                return FailResolve("locate_failed: 文档不可用");
            }

            if (target == null)
            {
                return new ResolveRangeResult
                {
                    Success = true,
                    UsedCursor = true,
                    Range = doc.Application.Selection.Range,
                    Spec = null
                };
            }

            try
            {
                if (target.Kind == LocatorKind.DocumentStart)
                {
                    Word.Range startRange = RangeAtDocumentStart(doc);
                    if (startRange == null)
                    {
                        return FailResolve("locate_failed: 无法取得文档开头 Range");
                    }

                    return OkRange(startRange, target, usedCursor: false);
                }

                if (refreshDocument && SpecNeedsProcessDocument(target))
                {
                    string pdError = TryProcessDocumentForLocate(doc, disallowBackendApi);
                    if (pdError != null)
                    {
                        return FailResolve("locate_failed: " + pdError);
                    }
                }

                if (target.Kind == LocatorKind.EmptyCell)
                {
                    return TryResolveEmptyCellRange(doc, target);
                }

                if (target.Kind == LocatorKind.Table)
                {
                    if (rangeKind == LocateRangeKind.ContentSpan)
                    {
                        return FailResolve("locate_failed: ContentSpan 仅支持句子/段落编码");
                    }

                    Word.Range afterTable = ObjectInsertRangeHelper.FindRangeAfterTable(doc, target.TableId);
                    if (afterTable == null)
                    {
                        return FailResolve($"locate_failed: 找不到表格 '{target.TableId}'");
                    }

                    return OkRange(afterTable, target, usedCursor: false);
                }

                if (target.Kind == LocatorKind.Chart)
                {
                    if (rangeKind == LocateRangeKind.ContentSpan)
                    {
                        return FailResolve("locate_failed: ContentSpan 仅支持句子/段落编码");
                    }

                    Word.Range afterChart = ObjectInsertRangeHelper.FindRangeAfterChart(doc, target.ChartId);
                    if (afterChart == null)
                    {
                        return FailResolve($"locate_failed: 找不到图表 '{target.ChartId}'");
                    }

                    return OkRange(afterChart, target, usedCursor: false);
                }

                if (target.Kind == LocatorKind.Image)
                {
                    if (rangeKind == LocateRangeKind.ContentSpan)
                    {
                        return FailResolve("locate_failed: ContentSpan 仅支持句子/段落编码");
                    }

                    Word.Range afterImage = ObjectInsertRangeHelper.FindRangeAfterImage(doc, target.ImageId);
                    if (afterImage == null)
                    {
                        return FailResolve($"locate_failed: 找不到图片 '{target.ImageId}'");
                    }

                    return OkRange(afterImage, target, usedCursor: false);
                }

                if (target.Kind == LocatorKind.Paragraph)
                {
                    return TryResolveSentenceOrParagraphRange(doc, target, rangeKind, paragraph: true);
                }

                if (target.Kind == LocatorKind.Sentence)
                {
                    return TryResolveSentenceOrParagraphRange(doc, target, rangeKind, paragraph: false);
                }

                return FailResolve("locate_failed: 无法识别的定位类型");
            }
            catch (Exception ex)
            {
                return FailResolve($"locate_failed: {ex.Message}");
            }
        }

        /// <summary>
        /// 整批歧义预检（停在歧义，不落 Range）。批里有号则一次 ProcessDocument。
        /// </summary>
        public static PrecheckResult TryPrecheckTargets(
            Word.Document doc,
            IReadOnlyList<DocumentActionAmbiguityAnchor> anchors,
            bool disallowBackendApi = false,
            LocatePrecheckStream stream = LocatePrecheckStream.SentenceActions,
            ApplyFormatScopeArgs formatScope = null,
            string formatToolName = null)
        {
            var result = new PrecheckResult { Success = true };
            if (doc == null)
            {
                result.Success = false;
                result.Error = "locate_failed: 文档不可用";
                return result;
            }

            if (AnchorsNeedProcessDocument(anchors, formatScope))
            {
                string pdError = TryProcessDocumentForLocate(doc, disallowBackendApi);
                if (pdError != null)
                {
                    result.Success = false;
                    result.Error = "locate_failed: " + pdError;
                    return result;
                }
            }

            string inoperable = FailIfInoperable(anchors);
            if (inoperable != null)
            {
                result.Success = false;
                result.Error = inoperable;
                return result;
            }

            string cacheError;
            AmbiguityCheckResult ambiguity;
            TableScopeIndex tableScope;
            if (stream == LocatePrecheckStream.ParagraphFormatApply
                || stream == LocatePrecheckStream.ParagraphFormatContext)
            {
                List<string> codes = FlattenAnchorCodes(anchors);
                if (stream == LocatePrecheckStream.ParagraphFormatContext)
                {
                    ambiguity = ParagraphFormatAmbiguityHelper.RunContextReadPrecheck(
                        codes,
                        formatScope ?? new ApplyFormatScopeArgs(),
                        formatToolName ?? "F_get_format_context",
                        out tableScope,
                        out cacheError);
                }
                else
                {
                    ambiguity = ParagraphFormatAmbiguityHelper.RunPrecheck(
                        formatScope ?? new ApplyFormatScopeArgs(),
                        codes,
                        formatToolName ?? "F_apply_paragraph_format",
                        out tableScope,
                        out cacheError);
                }
            }
            else if (stream == LocatePrecheckStream.SentenceFormatApply
                || stream == LocatePrecheckStream.SentenceFormatContext)
            {
                List<string> codes = FlattenAnchorCodes(anchors);
                if (stream == LocatePrecheckStream.SentenceFormatContext)
                {
                    ambiguity = ApplyFormatAmbiguityHelper.RunContextReadPrecheck(
                        codes,
                        formatScope ?? new ApplyFormatScopeArgs(),
                        formatToolName ?? "F_get_format_context",
                        out tableScope,
                        out cacheError);
                }
                else
                {
                    ambiguity = ApplyFormatAmbiguityHelper.RunPrecheck(
                        codes,
                        formatScope ?? new ApplyFormatScopeArgs(),
                        formatToolName ?? "F_apply_document_format",
                        out tableScope,
                        out cacheError);
                }
            }
            else
            {
                cacheError = null;
                SplitAnchorsByKind(anchors, out List<DocumentActionAmbiguityAnchor> sentenceAnchors,
                    out List<DocumentActionAmbiguityAnchor> paragraphAnchors);
                tableScope = TableSentenceScopeHelper.BuildIndexFromProcessCache(
                    DocumentState.LastProcessCache);
                AmbiguityCheckResult sentenceAmbiguity = sentenceAnchors.Count == 0
                    ? new AmbiguityCheckResult()
                    : SentenceCodeAmbiguityHelper.CheckBatch(
                        sentenceAnchors,
                        DocumentState.Snapshot,
                        ConfigManager.GetDocumentActionsAmbiguitySettings(),
                        tableScope);
                AmbiguityCheckResult paragraphAmbiguity = paragraphAnchors.Count == 0
                    ? new AmbiguityCheckResult()
                    : ParagraphCodeAmbiguityHelper.CheckBatch(
                        paragraphAnchors,
                        ConfigManager.GetDocumentActionsAmbiguitySettings(),
                        TableParagraphScopeHelper.BuildIndexFromProcessCache(DocumentState.LastProcessCache));
                if (paragraphAmbiguity != null && paragraphAmbiguity.HasBlockingError)
                {
                    ambiguity = paragraphAmbiguity;
                }
                else if (sentenceAmbiguity != null && sentenceAmbiguity.HasBlockingError)
                {
                    ambiguity = sentenceAmbiguity;
                }
                else
                {
                    ambiguity = MergeAmbiguity(sentenceAmbiguity, paragraphAmbiguity);
                }
            }

            result.TableScope = tableScope;
            result.Ambiguity = ambiguity;
            if (cacheError != null)
            {
                result.Success = false;
                result.Error = cacheError;
                return result;
            }

            if (ambiguity != null && ambiguity.HasBlockingError)
            {
                result.Success = false;
                result.Error = ambiguity.SummaryMessage ?? "locate_failed: 定位歧义";
                return result;
            }

            return result;
        }

        public static ResolveRangeResult TryResolveCodeSpan(
            Word.Document doc,
            IReadOnlyList<string> codes,
            string inTableId,
            LocateRangeKind rangeKind,
            bool paragraph,
            bool disallowBackendApi)
        {
            var spec = new LocatorSpec
            {
                Kind = paragraph ? LocatorKind.Paragraph : LocatorKind.Sentence,
                Codes = codes == null ? "" : string.Join(",", codes.Where(c => !string.IsNullOrWhiteSpace(c))),
                InTableId = inTableId
            };
            if (string.IsNullOrEmpty(spec.Codes))
            {
                return FailResolve("locate_failed: codes 为空");
            }

            if (doc == null)
            {
                return FailResolve("locate_failed: 文档不可用");
            }

            string pdError = TryProcessDocumentForLocate(doc, disallowBackendApi);
            if (pdError != null)
            {
                return FailResolve("locate_failed: " + pdError);
            }

            return TryResolveSentenceOrParagraphRange(doc, spec, rangeKind, paragraph);
        }

        public static ToolResult ToFailureResult(ResolveRangeResult locate, string toolName = null)
        {
            if (locate?.Ambiguity != null && locate.Ambiguity.HasBlockingError)
            {
                return SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                    string.IsNullOrEmpty(toolName) ? "locate" : toolName,
                    locate.Ambiguity);
            }

            return new ToolResult
            {
                Success = false,
                Error = locate?.Error ?? "locate_failed",
                Data = new Dictionary<string, object>
                {
                    ["error"] = "locate_failed",
                    ["checkpoint_cleanup_only"] = true
                }
            };
        }

        public static ToolResult ToPrecheckFailureResult(
            PrecheckResult precheck,
            string documentName,
            int totalActions)
        {
            if (precheck?.Error == StoryUniqEncode.FailBeforeWriteMessage)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = precheck.Error,
                    Data = new Dictionary<string, object>
                    {
                        ["checkpoint_cleanup_only"] = true
                    }
                };
            }

            if (precheck?.Ambiguity != null && precheck.Ambiguity.HasBlockingError)
            {
                return SentenceCodeAmbiguityHelper.BuildPreStep3FailureResult(
                    documentName,
                    totalActions,
                    precheck.Ambiguity);
            }

            return SentenceCodeAmbiguityHelper.BuildPreStep3FailureResult(
                documentName,
                totalActions,
                "locate_failed",
                precheck?.Error ?? "定位预检失败");
        }

        public static string TryProcessDocumentForLocate(Word.Document doc, bool disallowBackendApi)
        {
            try
            {
                using (EasyWriteDiagnostics.Time("process_actions.ProcessDocument_locate"))
                {
                    var options = ProcessDocumentOptions.ForProcessActions("locate");
                    options.DisallowBackendApi = disallowBackendApi;
                    WordDocumentExtractor.ProcessDocument(doc, options);
                }

                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static Dictionary<string, object> TryGetTargetDict(Dictionary<string, object> args)
        {
            return TryGetNamedLocatorDict(args, "target");
        }

        public static Dictionary<string, object> TryGetNamedLocatorDict(
            Dictionary<string, object> args,
            string key)
        {
            if (args == null || string.IsNullOrEmpty(key) || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            object raw = args[key];
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            Dictionary<string, object> coerced = ActionFormatHelper.CoerceDict(raw);
            if (coerced != null && coerced.Count > 0)
            {
                return coerced;
            }

            // JavaScriptSerializer 可能给出 Dictionary<string, object> 的变体
            if (raw is System.Collections.IDictionary idict)
            {
                var converted = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (System.Collections.DictionaryEntry entry in idict)
                {
                    if (entry.Key != null)
                    {
                        converted[entry.Key.ToString()] = entry.Value;
                    }
                }

                return converted.Count > 0 ? converted : null;
            }

            return null;
        }

        public static Dictionary<string, object> BuildTargetEcho(LocatorSpec spec)
        {
            if (spec == null)
            {
                return null;
            }

            var echo = new Dictionary<string, object>();
            if (spec.Kind == LocatorKind.Sentence)
            {
                echo["codes"] = spec.Codes;
                if (!string.IsNullOrEmpty(spec.InTableId))
                {
                    echo["in_table"] = spec.InTableId;
                }
            }
            else if (spec.Kind == LocatorKind.Paragraph)
            {
                echo["paragraph_codes"] = spec.Codes;
                if (!string.IsNullOrEmpty(spec.InTableId))
                {
                    echo["in_table"] = spec.InTableId;
                }
            }
            else if (spec.Kind == LocatorKind.Table)
            {
                echo["table"] = spec.TableId;
            }
            else if (spec.Kind == LocatorKind.Chart)
            {
                echo["chart"] = spec.ChartId;
            }
            else if (spec.Kind == LocatorKind.Image)
            {
                echo["image"] = spec.ImageId;
            }
            else if (spec.Kind == LocatorKind.DocumentStart)
            {
                echo["position"] = PositionStart;
            }

            return echo;
        }

        public static bool IsStartPosition(string raw)
        {
            return !string.IsNullOrWhiteSpace(raw)
                && string.Equals(raw.Trim(), PositionStart, StringComparison.OrdinalIgnoreCase);
        }

        public static Word.Range RangeAtDocumentStart(Word.Document doc)
        {
            if (doc == null)
            {
                return null;
            }

            int start = doc.Content.Start;
            return doc.Range(start, start);
        }

        public static string AppendDocumentStartMessage(string message, LocatorSpec spec)
        {
            if (spec == null || spec.Kind != LocatorKind.DocumentStart)
            {
                return message;
            }

            if (string.IsNullOrEmpty(message))
            {
                return "已插在正文最前。";
            }

            return message.TrimEnd() + " 已插在正文最前。";
        }

        private static ResolveRangeResult TryResolveSentenceOrParagraphRange(
            Word.Document doc,
            LocatorSpec target,
            LocateRangeKind rangeKind,
            bool paragraph)
        {
            var names = (target.Codes ?? "")
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
            if (names.Count == 0)
            {
                return FailResolve("locate_failed: target.codes 为空");
            }

            var formatScope = new ApplyFormatScopeArgs { TableId = target.InTableId };
            AmbiguityCheckResult ambiguity;
            TableScopeIndex tableScope;
            string cacheError;
            string field = paragraph ? "target_paragraph_codes" : "target";
            if (paragraph)
            {
                ambiguity = ParagraphFormatAmbiguityHelper.RunPrecheck(
                    formatScope,
                    names,
                    "locate",
                    out tableScope,
                    out cacheError);
            }
            else
            {
                var anchor = new DocumentActionAmbiguityAnchor
                {
                    ActionIndex = 0,
                    Field = field,
                    Codes = string.Join(",", names),
                    TableId = target.InTableId
                };
                tableScope = BuildSentenceTableScope(target.InTableId, out cacheError);
                ambiguity = SentenceCodeAmbiguityHelper.CheckBatch(
                    new[] { anchor },
                    DocumentState.Snapshot,
                    ConfigManager.GetDocumentActionsAmbiguitySettings(),
                    tableScope);
            }

            if (cacheError != null)
            {
                return FailResolve($"locate_failed: {cacheError}");
            }

            if (ambiguity != null && ambiguity.HasBlockingError)
            {
                return FailAmbiguity(ambiguity);
            }

            int displayStart = paragraph
                ? ParagraphFormatAmbiguityHelper.GetDisplayStartAt(ambiguity, 0, field)
                : GetSentenceDisplayStart(ambiguity, field);
            if (displayStart < 0)
            {
                return FailResolve("locate_failed: 无法定位编码序列");
            }

            if (!string.IsNullOrEmpty(target.InTableId)
                && (tableScope == null || !tableScope.ContainsTableId(target.InTableId)))
            {
                return FailResolve($"locate_failed: 表域 '{target.InTableId}' 不可用，请先 F_get_document_content");
            }

            HostProbe.Touch("sent-locate-doc", doc);
            HostProbe.Log(
                "sent-locate",
                "codes=" + string.Join(",", names)
                + " paragraph=" + paragraph
                + " displayStart=" + displayStart
                + " snap=" + (DocumentState.Snapshot == null ? "null" : "ok")
                + " stored=" + BriefStored(names));

            Word.Range range = paragraph
                ? DisplayPositionRangeResolver.ResolveParagraphSpan(
                    doc,
                    names,
                    displayStart,
                    target.InTableId,
                    tableScope,
                    debugTag: "locate_facade_para")
                : DisplayPositionRangeResolver.ResolveSpan(
                    doc,
                    names,
                    displayStart,
                    target.InTableId,
                    tableScope,
                    debugTag: "locate_facade");

            HostProbe.Log(
                "sent-range",
                range == null
                    ? "null"
                    : ("start=" + range.Start + " end=" + range.End
                        + " text=" + BriefText(range.Text)));

            if (range == null)
            {
                return FailResolve("locate_failed: 已解析编码但无法取得 Word Range");
            }

            if (rangeKind == LocateRangeKind.InsertAfter)
            {
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            return OkRange(range, target, usedCursor: false);
        }

        private static int GetSentenceDisplayStart(AmbiguityCheckResult ambiguity, string field)
        {
            if (ambiguity?.ResolvedDisplayStartPositionByField == null)
            {
                return -1;
            }

            string key = SentenceCodeAmbiguityHelper.MakeFieldKey(0, field);
            return ambiguity.ResolvedDisplayStartPositionByField.TryGetValue(key, out int start)
                ? start
                : -1;
        }

        private static bool SpecNeedsProcessDocument(LocatorSpec spec)
        {
            return spec != null
                && spec.Kind != LocatorKind.DocumentStart
                && (spec.Kind == LocatorKind.Sentence
                    || spec.Kind == LocatorKind.Paragraph
                    || spec.Kind == LocatorKind.Table
                    || spec.Kind == LocatorKind.Chart
                    || spec.Kind == LocatorKind.Image
                    || spec.Kind == LocatorKind.EmptyCell);
        }

        private static string FailIfInoperable(IReadOnlyList<DocumentActionAmbiguityAnchor> anchors)
        {
            if (anchors == null)
            {
                return null;
            }

            foreach (DocumentActionAmbiguityAnchor anchor in anchors)
            {
                if (StoryUniqEncode.CsvIsInoperable(anchor?.Codes))
                {
                    return StoryUniqEncode.FailBeforeWriteMessage;
                }
            }

            return null;
        }

        private static void SplitAnchorsByKind(
            IReadOnlyList<DocumentActionAmbiguityAnchor> anchors,
            out List<DocumentActionAmbiguityAnchor> sentenceAnchors,
            out List<DocumentActionAmbiguityAnchor> paragraphAnchors)
        {
            sentenceAnchors = new List<DocumentActionAmbiguityAnchor>();
            paragraphAnchors = new List<DocumentActionAmbiguityAnchor>();
            if (anchors == null)
            {
                return;
            }

            foreach (DocumentActionAmbiguityAnchor anchor in anchors)
            {
                if (AnchorLooksLikeParagraph(anchor))
                {
                    paragraphAnchors.Add(anchor);
                }
                else
                {
                    sentenceAnchors.Add(anchor);
                }
            }
        }

        private static bool AnchorLooksLikeParagraph(DocumentActionAmbiguityAnchor anchor)
        {
            if (string.IsNullOrWhiteSpace(anchor?.Codes))
            {
                return false;
            }

            string first = anchor.Codes.Split(',')[0].Trim();
            return FormatContextHelper.IsParagraphCode(first);
        }

        private static AmbiguityCheckResult MergeAmbiguity(
            AmbiguityCheckResult sentence,
            AmbiguityCheckResult paragraph)
        {
            var merged = new AmbiguityCheckResult();
            CopyAmbiguityMaps(sentence, merged);
            CopyAmbiguityMaps(paragraph, merged);
            return merged;
        }

        private static void CopyAmbiguityMaps(AmbiguityCheckResult source, AmbiguityCheckResult dest)
        {
            if (source?.ResolvedDisplayStartPositionByField == null)
            {
                return;
            }

            foreach (var pair in source.ResolvedDisplayStartPositionByField)
            {
                dest.ResolvedDisplayStartPositionByField[pair.Key] = pair.Value;
            }

            if (source.ResolvedOccurrenceIndexesByField == null)
            {
                return;
            }

            foreach (var pair in source.ResolvedOccurrenceIndexesByField)
            {
                dest.ResolvedOccurrenceIndexesByField[pair.Key] = pair.Value;
            }
        }

        private static bool AnchorsNeedProcessDocument(
            IReadOnlyList<DocumentActionAmbiguityAnchor> anchors,
            ApplyFormatScopeArgs formatScope)
        {
            if (!string.IsNullOrWhiteSpace(formatScope?.TableId))
            {
                return true;
            }

            if (anchors == null)
            {
                return false;
            }

            foreach (DocumentActionAmbiguityAnchor anchor in anchors)
            {
                if (!string.IsNullOrWhiteSpace(anchor?.Codes)
                    || !string.IsNullOrWhiteSpace(anchor?.TableId))
                {
                    return true;
                }
            }

            return false;
        }

        private static List<string> FlattenAnchorCodes(IReadOnlyList<DocumentActionAmbiguityAnchor> anchors)
        {
            var codes = new List<string>();
            if (anchors == null)
            {
                return codes;
            }

            foreach (DocumentActionAmbiguityAnchor anchor in anchors)
            {
                if (string.IsNullOrWhiteSpace(anchor?.Codes))
                {
                    continue;
                }

                foreach (string part in anchor.Codes.Split(','))
                {
                    string trimmed = part.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                    {
                        codes.Add(trimmed);
                    }
                }
            }

            return codes;
        }

        private static TableScopeIndex BuildSentenceTableScope(string tableId, out string cacheError)
        {
            cacheError = null;
            if (string.IsNullOrEmpty(tableId))
            {
                return TableSentenceScopeHelper.BuildIndexFromProcessCache(DocumentState.LastProcessCache);
            }

            var scopeArgs = new ApplyFormatScopeArgs { TableId = tableId };
            return ApplyFormatAmbiguityHelper.BuildTableScopeOrNull(scopeArgs, out cacheError);
        }

        private static ResolveRangeResult FailAmbiguity(AmbiguityCheckResult ambiguity)
        {
            HostProbe.Log("sent-fail", ambiguity?.SummaryMessage ?? "ambiguous");
            return new ResolveRangeResult
            {
                Success = false,
                Error = ambiguity?.SummaryMessage
                    ?? $"locate_failed: 句子编码歧义（{DisplaySequenceLocator.ErrorAmbiguousSentenceCode}）",
                Ambiguity = ambiguity
            };
        }

        private static ResolveRangeResult TryResolveEmptyCellRange(Word.Document doc, LocatorSpec target)
        {
            if (target == null || string.IsNullOrEmpty(target.TableId) || target.Row < 1 || target.Col < 1)
            {
                return FailResolve("locate_failed: 空格码缺少表号或行列");
            }

            Word.Table table = TableResolveHelper.TryResolveTableById(doc, target.TableId, out string tableError);
            if (table == null)
            {
                return FailResolve("locate_failed: " + (tableError ?? $"找不到表格 '{target.TableId}'"));
            }

            Word.Cell cell = TableRowValuesHelper.TryResolveWritableCell(table, target.Row, target.Col);
            if (cell == null)
            {
                return FailResolve(
                    "locate_failed: " + TableRowValuesHelper.BuildUnresolvableCellError(target.Row, target.Col));
            }

            return OkRange(cell.Range, target, usedCursor: false);
        }

        private static ResolveRangeResult OkRange(Word.Range range, LocatorSpec spec, bool usedCursor)
        {
            return new ResolveRangeResult
            {
                Success = true,
                Range = range,
                Spec = spec,
                UsedCursor = usedCursor
            };
        }

        private static ResolveRangeResult FailResolve(string error)
        {
            HostProbe.Log("sent-fail", error ?? "");
            return new ResolveRangeResult { Success = false, Error = error };
        }

        private static string BriefStored(IList<string> names)
        {
            if (names == null || names.Count == 0)
            {
                return "(none)";
            }

            var parts = new List<string>();
            foreach (string name in names)
            {
                string text = DocumentState.GetSentenceContent(name);
                parts.Add(name + "=" + BriefText(text));
            }

            return string.Join(" | ", parts);
        }

        private static string BriefText(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            string escaped = TextUtils.EscapeNewlinesForDb(text);
            return escaped.Length <= 160 ? escaped : escaped.Substring(0, 160);
        }
    }
}
