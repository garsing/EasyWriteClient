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
            public TextSpan TextSpan { get; set; }
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

        /// <summary>一批里的一项。ActionIndex、Field 只供预检把失败填回调用方原来的槽。</summary>
        public sealed class LocateInput
        {
            public Dictionary<string, object> Locator { get; set; }
            public int ActionIndex { get; set; } = -1;
            public string Field { get; set; }
            public string InsertContentPlaceholder { get; set; }
        }

        /// <summary>与入参下标对齐的一项。对不上时 Range 为空。</summary>
        public sealed class LocateSpanItem
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public Word.Range Range { get; set; }
            public LocatorSpec Spec { get; set; }
            public Word.Table Table { get; set; }
            public Word.InlineShape Image { get; set; }
            public ChartResolveHelper.BodyChart Chart { get; set; }
        }

        public sealed class LocateSpanBatchResult
        {
            public bool BatchBlocked { get; set; }
            public string Error { get; set; }
            public List<LocateSpanItem> Items { get; set; } = new List<LocateSpanItem>();
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

                return WithText(
                    new LocatorParseResult
                    {
                        Success = true,
                        Spec = new LocatorSpec { Kind = LocatorKind.DocumentStart }
                    },
                    dict,
                    singleCode: false);
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

                return WithText(
                    new LocatorParseResult
                    {
                        Success = true,
                        Spec = new LocatorSpec { Kind = LocatorKind.Table, TableId = tableId }
                    },
                    dict,
                    singleCode: false);
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

                return WithText(
                    new LocatorParseResult
                    {
                        Success = true,
                        Spec = new LocatorSpec { Kind = LocatorKind.Chart, ChartId = chartId }
                    },
                    dict,
                    singleCode: false);
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

            return WithText(
                new LocatorParseResult
                {
                    Success = true,
                    Spec = new LocatorSpec { Kind = LocatorKind.Image, ImageId = imageId }
                },
                dict,
                singleCode: false);
        }

        private static LocatorParseResult FailParse(string error)
        {
            return new LocatorParseResult { Success = false, Error = error };
        }

        private static LocatorParseResult WithText(
            LocatorParseResult parsed,
            Dictionary<string, object> dict,
            bool singleCode)
        {
            if (parsed == null || !parsed.Success)
            {
                return parsed;
            }

            if (!TextSpanArg.TryRead(dict, out TextSpan span, out string error))
            {
                return FailParse(error);
            }

            if (span == null)
            {
                return parsed;
            }

            if (!singleCode)
            {
                return FailParse(TextSpanArg.SingleCodeError);
            }

            parsed.Spec.TextSpan = span;
            return parsed;
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

                return WithText(
                    new LocatorParseResult
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
                    },
                    dict,
                    singleCode: true);
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
                    if (EmptyCellCode.IsIllegalCharsetGeometricCode(part))
                    {
                        return FailParse($"{fieldLabel}.codes 表号含非字母数字，几何码无法写入");
                    }

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

            return WithText(
                new LocatorParseResult
                {
                    Success = true,
                    Spec = new LocatorSpec
                    {
                        Kind = LocatorKind.Sentence,
                        Codes = string.Join(",", parts),
                        InTableId = inTable
                    }
                },
                dict,
                singleCode: parts.Count == 1);
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

            return WithText(
                new LocatorParseResult
                {
                    Success = true,
                    Spec = new LocatorSpec
                    {
                        Kind = LocatorKind.Paragraph,
                        Codes = string.Join(",", parts),
                        InTableId = inTable
                    }
                },
                dict,
                singleCode: parts.Count == 1);
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
                if (TextSpanArg.HasNarrowingText(args))
                {
                    return FailResolve(TextSpanArg.SingleCodeError);
                }

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
                string parsedError = parsed.Error ?? "定位失败";
                if (TextSpanArg.IsDirectError(parsedError))
                {
                    return FailResolve(parsedError);
                }

                return FailResolve($"locate_failed: {parsedError}");
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
                if (target.TextSpan != null
                    && target.Kind != LocatorKind.Sentence
                    && target.Kind != LocatorKind.Paragraph)
                {
                    return FailResolve(TextSpanArg.SingleCodeError);
                }

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

                    Word.Table tableObj = SentenceCodeLocator.ResolveTableById(doc, target.TableId);
                    if (tableObj == null)
                    {
                        return FailResolve($"locate_failed: 找不到表格 '{target.TableId}'");
                    }

                    Word.Range afterTable = ObjectInsertRangeHelper.FindRangeAfterTable(doc, target.TableId);
                    if (afterTable == null)
                    {
                        return FailResolve(ObjectInsertRangeHelper.TableIsNested(tableObj, doc)
                            ? $"locate_failed: 无法在嵌套表 '{target.TableId}' 后插入（需落在父单元格内）"
                            : $"locate_failed: 无法在表格 '{target.TableId}' 后插入");
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
                LocateProbeLog.Line(
                    "precheck.block stream=" + stream
                    + " err=" + inoperable
                    + " " + TocFieldStamp.DumpInterior());
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
                ambiguity = MergeAmbiguity(sentenceAmbiguity, paragraphAmbiguity);
                if (paragraphAmbiguity != null && paragraphAmbiguity.HasBlockingError)
                {
                    ambiguity.HasBlockingError = true;
                    ambiguity.ErrorCode = paragraphAmbiguity.ErrorCode;
                    ambiguity.SummaryMessage = paragraphAmbiguity.SummaryMessage;
                    ambiguity.Hint = paragraphAmbiguity.Hint;
                    ambiguity.AmbiguousCodes = paragraphAmbiguity.AmbiguousCodes
                        ?? new List<Dictionary<string, object>>();
                }
                else if (sentenceAmbiguity != null && sentenceAmbiguity.HasBlockingError)
                {
                    ambiguity.HasBlockingError = true;
                    ambiguity.ErrorCode = sentenceAmbiguity.ErrorCode;
                    ambiguity.SummaryMessage = sentenceAmbiguity.SummaryMessage;
                    ambiguity.Hint = sentenceAmbiguity.Hint;
                    ambiguity.AmbiguousCodes = sentenceAmbiguity.AmbiguousCodes
                        ?? new List<Dictionary<string, object>>();
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

        /// <summary>
        /// 统一定位入口。解析整批，句子/段落/几何码一次预检，再交整段 Range。
        /// 表、图表、图片和文首不进这次预检。
        /// </summary>
        public static LocateSpanBatchResult TryLocateSpans(
            Word.Document doc,
            IReadOnlyList<Dictionary<string, object>> locators,
            bool disallowBackendApi = false,
            LocatePrecheckStream stream = LocatePrecheckStream.SentenceActions,
            ApplyFormatScopeArgs formatScope = null,
            string formatToolName = null)
        {
            var inputs = new List<LocateInput>();
            if (locators != null)
            {
                for (int i = 0; i < locators.Count; i++)
                {
                    inputs.Add(new LocateInput
                    {
                        Locator = locators[i],
                        ActionIndex = i,
                        Field = "item"
                    });
                }
            }

            return TryLocateSpans(doc, inputs, disallowBackendApi, stream, formatScope, formatToolName);
        }

        public static LocateSpanBatchResult LocateFormatSpan(
            Word.Document doc,
            IReadOnlyList<string> codes,
            bool paragraph,
            string tableId,
            TextSpan textSpan,
            bool disallowBackendApi,
            LocatePrecheckStream stream,
            ApplyFormatScopeArgs formatScope,
            string formatToolName)
        {
            var locator = new Dictionary<string, object>
            {
                [paragraph ? "paragraph_codes" : "codes"] = codes == null ? "" : string.Join(",", codes)
            };
            if (!string.IsNullOrEmpty(tableId))
            {
                locator["in_table"] = tableId;
            }

            if (textSpan != null && !string.IsNullOrEmpty(textSpan.Value))
            {
                locator["text"] = textSpan.Value;
            }

            return TryLocateSpans(
                doc,
                new[] { locator },
                disallowBackendApi,
                stream,
                formatScope,
                formatToolName);
        }

        public static LocateSpanBatchResult TryLocateSpans(
            Word.Document doc,
            IReadOnlyList<LocateInput> inputs,
            bool disallowBackendApi = false,
            LocatePrecheckStream stream = LocatePrecheckStream.SentenceActions,
            ApplyFormatScopeArgs formatScope = null,
            string formatToolName = null)
        {
            var result = new LocateSpanBatchResult();
            int count = inputs == null ? 0 : inputs.Count;
            for (int i = 0; i < count; i++)
            {
                result.Items.Add(new LocateSpanItem());
            }

            if (doc == null)
            {
                result.BatchBlocked = true;
                result.Error = "locate_failed: 文档不可用";
                StampBatch(result, result.Error);
                return result;
            }

            var anchors = new List<DocumentActionAmbiguityAnchor>();
            bool needRead = false;
            for (int i = 0; i < count; i++)
            {
                LocateInput input = inputs[i];
                LocatorParseResult parsed = ParseLocatorObject(
                    input?.Locator,
                    sentenceOnly: false,
                    fieldLabel: string.IsNullOrEmpty(input?.Field) ? "target" : input.Field);
                if (parsed == null || !parsed.Success || parsed.Spec == null)
                {
                    result.Items[i].Success = false;
                    result.Items[i].Error = parsed?.Error ?? "定位对象无效";
                    continue;
                }

                result.Items[i].Spec = parsed.Spec;
                if (SpecNeedsProcessDocument(parsed.Spec))
                {
                    needRead = true;
                }

                if (parsed.Spec.Kind != LocatorKind.Sentence
                    && parsed.Spec.Kind != LocatorKind.Paragraph
                    && parsed.Spec.Kind != LocatorKind.EmptyCell)
                {
                    continue;
                }

                int actionIndex = input.ActionIndex >= 0 ? input.ActionIndex : i;
                anchors.Add(new DocumentActionAmbiguityAnchor
                {
                    ActionIndex = actionIndex,
                    Field = string.IsNullOrEmpty(input.Field) ? "item" : input.Field,
                    Codes = parsed.Spec.Codes,
                    TableId = !string.IsNullOrEmpty(parsed.Spec.InTableId)
                        ? parsed.Spec.InTableId
                        : (formatScope == null ? null : formatScope.TableId),
                    InsertContentPlaceholder = input.InsertContentPlaceholder
                });
            }

            if (needRead)
            {
                string pdError = TryProcessDocumentForLocate(doc, disallowBackendApi);
                if (pdError != null)
                {
                    result.BatchBlocked = true;
                    result.Error = "locate_failed: " + pdError;
                    StampBatch(result, result.Error);
                    return result;
                }
            }

            PrecheckResult precheck = null;
            if (anchors.Count > 0)
            {
                precheck = TryPrecheckTargets(
                    doc,
                    anchors,
                    disallowBackendApi,
                    stream,
                    formatScope,
                    formatToolName);
                result.Ambiguity = precheck.Ambiguity;
                result.TableScope = precheck.TableScope;
            }

            bool split = precheck?.Ambiguity?.LocateFailures != null
                && precheck.Ambiguity.LocateFailures.Count > 0;
            if (precheck != null && !precheck.Success && !split)
            {
                result.BatchBlocked = true;
                result.Error = precheck.Error ?? "定位预检失败";
                StampBatch(result, result.Error);
                return result;
            }

            var failedIndexes = new HashSet<int>();
            if (split)
            {
                foreach (AnchorLocateFailure failure in precheck.Ambiguity.LocateFailures)
                {
                    if (failure == null)
                    {
                        continue;
                    }

                    failedIndexes.Add(failure.ActionIndex);
                    for (int i = 0; i < count; i++)
                    {
                        int actionIndex = inputs[i].ActionIndex >= 0 ? inputs[i].ActionIndex : i;
                        if (actionIndex != failure.ActionIndex)
                        {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(inputs[i].Field)
                            && !string.Equals(inputs[i].Field, failure.Field, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        result.Items[i].Success = false;
                        result.Items[i].Error = string.IsNullOrEmpty(failure.Message)
                            ? (precheck.Error ?? "定位失败")
                            : failure.Message;
                    }
                }
            }

            for (int i = 0; i < count; i++)
            {
                LocateSpanItem item = result.Items[i];
                if (item.Spec == null || !string.IsNullOrEmpty(item.Error))
                {
                    item.Success = false;
                    continue;
                }

                int actionIndex = inputs[i].ActionIndex >= 0 ? inputs[i].ActionIndex : i;
                if (failedIndexes.Contains(actionIndex)
                    && (string.IsNullOrEmpty(inputs[i].Field) || ItemMatchesFailureField(inputs[i], precheck)))
                {
                    item.Success = false;
                    if (string.IsNullOrEmpty(item.Error))
                    {
                        item.Error = "定位失败";
                    }

                    continue;
                }

                FillWholeSpan(doc, item);
            }

            return result;
        }

        private static bool ItemMatchesFailureField(LocateInput input, PrecheckResult precheck)
        {
            if (input == null || precheck?.Ambiguity?.LocateFailures == null)
            {
                return true;
            }

            int actionIndex = input.ActionIndex;
            foreach (AnchorLocateFailure failure in precheck.Ambiguity.LocateFailures)
            {
                if (failure == null || failure.ActionIndex != actionIndex)
                {
                    continue;
                }

                if (string.IsNullOrEmpty(input.Field)
                    || string.Equals(input.Field, failure.Field, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static void StampBatch(LocateSpanBatchResult result, string error)
        {
            if (result?.Items == null)
            {
                return;
            }

            foreach (LocateSpanItem item in result.Items)
            {
                item.Success = false;
                item.Range = null;
                item.Table = null;
                item.Image = null;
                item.Chart = null;
                if (string.IsNullOrEmpty(item.Error))
                {
                    item.Error = error;
                }
            }
        }

        private static void FillWholeSpan(Word.Document doc, LocateSpanItem item)
        {
            LocatorSpec spec = item.Spec;
            try
            {
                if (spec.Kind == LocatorKind.DocumentStart)
                {
                    Word.Range startRange = RangeAtDocumentStart(doc);
                    if (startRange == null)
                    {
                        item.Success = false;
                        item.Error = "locate_failed: 无法取得文档开头 Range";
                        return;
                    }

                    item.Range = startRange;
                    item.Success = true;
                    return;
                }

                if (spec.Kind == LocatorKind.Sentence || spec.Kind == LocatorKind.Paragraph)
                {
                    ResolveRangeResult located = TryResolveSentenceOrParagraphRange(
                        doc,
                        spec,
                        LocateRangeKind.ContentSpan,
                        spec.Kind == LocatorKind.Paragraph);
                    item.Success = located.Success && located.Range != null;
                    item.Range = located.Range;
                    item.Error = item.Success ? null : (located.Error ?? "locate_failed: 无法定位编码序列");
                    return;
                }

                if (spec.Kind == LocatorKind.EmptyCell)
                {
                    ResolveRangeResult located = TryResolveEmptyCellRange(doc, spec);
                    item.Success = located.Success && located.Range != null;
                    item.Range = located.Range;
                    item.Error = item.Success ? null : (located.Error ?? "locate_failed: 无法定位空格");
                    if (item.Success)
                    {
                        item.Table = TableResolveHelper.TryResolveTableById(doc, spec.TableId, out _);
                        if (spec.TextSpan != null && !string.IsNullOrEmpty(spec.TextSpan.Value))
                        {
                            ResolveRangeResult narrowed = TextSpanNarrower.TryNarrow(
                                doc,
                                item.Range,
                                spec.TextSpan,
                                spec.Codes);
                            if (narrowed == null || !narrowed.Success || narrowed.Range == null)
                            {
                                item.Success = false;
                                item.Range = null;
                                item.Error = narrowed == null || string.IsNullOrEmpty(narrowed.Error)
                                    ? "locate_failed: 无法定位空格"
                                    : narrowed.Error;
                                return;
                            }

                            item.Range = narrowed.Range;
                        }
                    }

                    return;
                }

                if (spec.Kind == LocatorKind.Table)
                {
                    Word.Table table = TableResolveHelper.TryResolveTableById(doc, spec.TableId, out string tableError);
                    if (table == null)
                    {
                        item.Success = false;
                        item.Error = "locate_failed: " + (tableError ?? $"找不到表格 '{spec.TableId}'");
                        return;
                    }

                    item.Table = table;
                    item.Range = table.Range;
                    item.Success = item.Range != null;
                    if (!item.Success)
                    {
                        item.Error = $"locate_failed: 无法取得表格 '{spec.TableId}' 的 Range";
                    }

                    return;
                }

                if (spec.Kind == LocatorKind.Chart)
                {
                    ChartResolveHelper.BodyChart chart = ChartResolveHelper.ResolveBodyChartByChartId(doc, spec.ChartId);
                    if (chart == null)
                    {
                        item.Success = false;
                        item.Error = $"locate_failed: 找不到图表 '{spec.ChartId}'";
                        return;
                    }

                    item.Chart = chart;
                    item.Range = chart.NavigateRange();
                    item.Success = item.Range != null;
                    if (!item.Success)
                    {
                        item.Error = $"locate_failed: 无法取得图表 '{spec.ChartId}' 的 Range";
                    }

                    return;
                }

                if (spec.Kind == LocatorKind.Image)
                {
                    Word.InlineShape image = ImageResolveHelper.ResolveInlineShapeByImageId(doc, spec.ImageId);
                    if (image == null)
                    {
                        item.Success = false;
                        item.Error = $"locate_failed: 找不到图片 '{spec.ImageId}'";
                        return;
                    }

                    item.Image = image;
                    item.Range = image.Range;
                    item.Success = item.Range != null;
                    if (!item.Success)
                    {
                        item.Error = $"locate_failed: 无法取得图片 '{spec.ImageId}' 的 Range";
                    }

                    return;
                }

                item.Success = false;
                item.Error = "locate_failed: 无法识别的定位类型";
            }
            catch (Exception ex)
            {
                item.Success = false;
                item.Range = null;
                item.Error = "locate_failed: " + ex.Message;
            }
        }

        /// <summary>
        /// 插入类调用方：没有 target 时用当前光标；有 target 时先取整段，再落到插入点。
        /// </summary>
        public static bool TryTakeInsertRange(
            Word.Document doc,
            Dictionary<string, object> args,
            bool disallowBackendApi,
            out Word.Range range,
            out string error,
            out bool usedCursor,
            out LocatorSpec spec)
        {
            range = null;
            error = null;
            usedCursor = false;
            spec = null;
            Dictionary<string, object> target = TryGetTargetDict(args);
            if (target == null)
            {
                if (doc == null)
                {
                    error = "locate_failed: 文档不可用";
                    return false;
                }

                try
                {
                    range = doc.Application.Selection.Range;
                    usedCursor = true;
                    return range != null || FailInsertPoint(out error, "没有活动选区");
                }
                catch (Exception ex)
                {
                    error = "locate_failed: " + ex.Message;
                    return false;
                }
            }

            LocateSpanBatchResult batch = TryLocateSpans(
                doc,
                new[] { target },
                disallowBackendApi);
            LocateSpanItem item = batch.Items.Count > 0 ? batch.Items[0] : null;
            if (batch.BatchBlocked || item == null || !item.Success)
            {
                error = item?.Error ?? batch.Error ?? "无法确定插入位置";
                return false;
            }

            spec = item.Spec;
            return TryInsertPoint(doc, item, out range, out error);
        }

        /// <summary>调用方要插在整段后面时使用。文首和光标不在这里处理。</summary>
        public static bool TryInsertPoint(
            Word.Document doc,
            LocateSpanItem item,
            out Word.Range range,
            out string error)
        {
            range = null;
            error = null;
            if (item == null || !item.Success || item.Range == null || item.Spec == null)
            {
                error = item?.Error ?? "无法确定插入位置";
                return false;
            }

            try
            {
                if (item.Spec.Kind == LocatorKind.Table)
                {
                    range = ObjectInsertRangeHelper.InsertPointAfterTable(
                        doc, item.Table, item.Range, item.Spec.TableId);
                    if (range == null)
                    {
                        error = ObjectInsertRangeHelper.TableIsNested(item.Table, doc)
                            ? $"locate_failed: 无法在嵌套表 '{item.Spec.TableId}' 后插入（需落在父单元格内）"
                            : $"locate_failed: 无法在表格 '{item.Spec.TableId}' 后插入";
                        return false;
                    }

                    return true;
                }

                if (item.Spec.Kind == LocatorKind.Chart || item.Spec.Kind == LocatorKind.Image)
                {
                    range = ObjectInsertRangeHelper.InsertPointAfterShape(item.Range);
                    return range != null || FailInsertPoint(out error, "无法在对象后插入");
                }

                if (item.Spec.Kind == LocatorKind.DocumentStart)
                {
                    range = item.Range;
                    return true;
                }

                range = ObjectInsertRangeHelper.PointAtEnd(item.Range);
                return range != null || FailInsertPoint(out error, "无法确定插入位置");
            }
            catch (Exception ex)
            {
                error = "locate_failed: " + ex.Message;
                return false;
            }
        }

        private static bool FailInsertPoint(out string error, string message)
        {
            error = message;
            return false;
        }

        public static ResolveRangeResult TryResolveCodeSpan(
            Word.Document doc,
            IReadOnlyList<string> codes,
            string inTableId,
            LocateRangeKind rangeKind,
            bool paragraph,
            bool disallowBackendApi,
            TextSpan textSpan = null)
        {
            var spec = new LocatorSpec
            {
                Kind = paragraph ? LocatorKind.Paragraph : LocatorKind.Sentence,
                Codes = codes == null ? "" : string.Join(",", codes.Where(c => !string.IsNullOrWhiteSpace(c))),
                InTableId = inTableId,
                TextSpan = textSpan
            };
            if (string.IsNullOrEmpty(spec.Codes))
            {
                return FailResolve("locate_failed: codes 为空");
            }

            if (textSpan != null && spec.Codes.IndexOf(',') >= 0)
            {
                return FailResolve(TextSpanArg.SingleCodeError);
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
            else if (spec.Kind == LocatorKind.EmptyCell)
            {
                echo["codes"] = spec.Codes;
                if (!string.IsNullOrEmpty(spec.InTableId))
                {
                    echo["in_table"] = spec.InTableId;
                }
            }

            if (spec.TextSpan != null && !string.IsNullOrEmpty(spec.TextSpan.Value))
            {
                echo["text"] = spec.TextSpan.Value;
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

            if (target.TextSpan != null)
            {
                string[] codeParts = target.Codes.Split(',');
                int codeCount = 0;
                string onlyCode = null;
                foreach (string part in codeParts)
                {
                    if (string.IsNullOrWhiteSpace(part))
                    {
                        continue;
                    }

                    codeCount++;
                    onlyCode = part.Trim();
                }

                if (codeCount != 1)
                {
                    return FailResolve(TextSpanArg.SingleCodeError);
                }

                ResolveRangeResult narrowed = TextSpanNarrower.TryNarrow(doc, range, target.TextSpan, onlyCode);
                if (!narrowed.Success || narrowed.Range == null)
                {
                    return FailResolve(narrowed.Error ?? TextSpanArg.Missing(onlyCode, target.TextSpan.Value));
                }

                range = narrowed.Range;
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
                string codes = anchor?.Codes;
                bool inop = StoryUniqEncode.CsvIsInoperable(codes);
                bool interior = TocFieldStamp.IsInterior(codes);
                LocateProbeLog.Line(
                    "FailIfInoperable idx=" + (anchor == null ? -1 : anchor.ActionIndex)
                    + " field=" + (anchor == null ? "" : anchor.Field)
                    + " codes=" + (codes ?? "")
                    + " inop=" + inop
                    + " interior=" + interior
                    + " " + TocFieldStamp.DumpInterior());
                if (inop)
                {
                    return StoryUniqEncode.FailBeforeWriteMessage;
                }

                if (interior)
                {
                    return TocFieldStamp.InteriorMessage;
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

            if (source.LocateFailures == null)
            {
                return;
            }

            if (dest.LocateFailures == null)
            {
                dest.LocateFailures = new List<AnchorLocateFailure>();
            }

            dest.LocateFailures.AddRange(source.LocateFailures);
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
