using System;
using System.Collections.Generic;
using System.Linq;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// insert / F_insert_svg_image 共用的 target oneof 解析与落点 Range。
    /// 解析规则与 F_ProcessDocumentActionsTool 历史实现一致。
    /// </summary>
    public static class InsertTargetLocatorHelper
    {
        public enum LocatorKind
        {
            Sentence,
            Table,
            Chart,
            Image
        }

        public sealed class LocatorSpec
        {
            public LocatorKind Kind { get; set; }
            public string Codes { get; set; }
            public string InTableId { get; set; }
            public string TableId { get; set; }
            public string ChartId { get; set; }
            public string ImageId { get; set; }
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
        }

        /// <summary>
        /// 解析定位对象：codes(+in_table) / table / chart / image 四选一。
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

            bool hasCodes = dict.ContainsKey("codes") && dict["codes"] != null
                && !string.IsNullOrWhiteSpace(dict["codes"].ToString());
            bool hasTable = dict.ContainsKey("table") && dict["table"] != null
                && !string.IsNullOrWhiteSpace(dict["table"].ToString());
            bool hasChart = dict.ContainsKey("chart") && dict["chart"] != null
                && !string.IsNullOrWhiteSpace(dict["chart"].ToString());
            bool hasImage = dict.ContainsKey("image") && dict["image"] != null
                && !string.IsNullOrWhiteSpace(dict["image"].ToString());
            bool hasInTable = dict.ContainsKey("in_table") && dict["in_table"] != null
                && !string.IsNullOrWhiteSpace(dict["in_table"].ToString());

            int modeCount = (hasCodes ? 1 : 0) + (hasTable ? 1 : 0) + (hasChart ? 1 : 0) + (hasImage ? 1 : 0);
            if (modeCount != 1)
            {
                return new LocatorParseResult
                {
                    Success = false,
                    Error = $"{fieldLabel} 定位对象须为 codes / table / chart / image 四选一"
                };
            }

            if (sentenceOnly && (hasTable || hasChart || hasImage))
            {
                return new LocatorParseResult
                {
                    Success = false,
                    Error = $"{fieldLabel} 仅支持句子定位（codes + 可选 in_table）"
                };
            }

            if (hasCodes)
            {
                string codes = dict["codes"].ToString().Trim();
                var parts = codes.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
                if (parts.Count == 0)
                {
                    return new LocatorParseResult { Success = false, Error = $"{fieldLabel}.codes 为空" };
                }

                foreach (string part in parts)
                {
                    if (part.StartsWith("T_", StringComparison.Ordinal)
                        || part.StartsWith("C_", StringComparison.Ordinal)
                        || part.StartsWith("I_", StringComparison.Ordinal))
                    {
                        return new LocatorParseResult
                        {
                            Success = false,
                            Error = $"{fieldLabel}.codes 仅支持 S_；表/图/图片请用 table / chart / image 字段"
                        };
                    }

                    if (!part.StartsWith("S_", StringComparison.Ordinal))
                    {
                        return new LocatorParseResult
                        {
                            Success = false,
                            Error = $"{fieldLabel}.codes 仅支持 S_ 句子编码"
                        };
                    }
                }

                string inTable = hasInTable ? dict["in_table"].ToString().Trim() : null;
                if (!string.IsNullOrEmpty(inTable) && !inTable.StartsWith("T_", StringComparison.Ordinal))
                {
                    return new LocatorParseResult
                    {
                        Success = false,
                        Error = $"{fieldLabel}.in_table 须为 T_ 表格编号"
                    };
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

            if (hasInTable)
            {
                return new LocatorParseResult
                {
                    Success = false,
                    Error = $"{fieldLabel}.in_table 仅可与 codes 同用"
                };
            }

            if (hasTable)
            {
                string tableId = dict["table"].ToString().Trim();
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

        /// <summary>
        /// 解析 args["target"]（可省略）。省略则光标处。
        /// </summary>
        public static ResolveRangeResult TryResolveInsertRange(
            Word.Document doc,
            Dictionary<string, object> args)
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

            return TryResolveInsertRange(doc, parsed.Spec);
        }

        public static ResolveRangeResult TryResolveInsertRange(Word.Document doc, LocatorSpec target)
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
                if (target.Kind == LocatorKind.Table)
                {
                    Word.Range afterTable = ObjectInsertRangeHelper.FindRangeAfterTable(doc, target.TableId);
                    if (afterTable == null)
                    {
                        return FailResolve($"locate_failed: 找不到表格 '{target.TableId}'");
                    }

                    return OkRange(afterTable, target, usedCursor: false);
                }

                if (target.Kind == LocatorKind.Chart)
                {
                    Word.Range afterChart = ObjectInsertRangeHelper.FindRangeAfterChart(doc, target.ChartId);
                    if (afterChart == null)
                    {
                        return FailResolve($"locate_failed: 找不到图表 '{target.ChartId}'");
                    }

                    return OkRange(afterChart, target, usedCursor: false);
                }

                if (target.Kind == LocatorKind.Image)
                {
                    Word.Range afterImage = ObjectInsertRangeHelper.FindRangeAfterImage(doc, target.ImageId);
                    if (afterImage == null)
                    {
                        return FailResolve($"locate_failed: 找不到图片 '{target.ImageId}'");
                    }

                    return OkRange(afterImage, target, usedCursor: false);
                }

                // Sentence
                var names = target.Codes.Split(',')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToList();
                if (names.Count == 0)
                {
                    return FailResolve("locate_failed: target.codes 为空");
                }

                TableScopeIndex tableScope = null;
                if (!string.IsNullOrEmpty(target.InTableId))
                {
                    var scopeArgs = new ApplyFormatScopeArgs { TableId = target.InTableId };
                    tableScope = ApplyFormatAmbiguityHelper.BuildTableScopeOrNull(scopeArgs, out string cacheError);
                    if (cacheError != null)
                    {
                        return FailResolve($"locate_failed: {cacheError}");
                    }

                    if (tableScope == null || !tableScope.ContainsTableId(target.InTableId))
                    {
                        return FailResolve($"locate_failed: 表域 '{target.InTableId}' 不可用，请先 F_get_document_content");
                    }
                }

                var locateOptions = new DisplaySequenceLocateOptions
                {
                    TableId = target.InTableId,
                    TableScope = tableScope,
                    StreamKind = CodeStreamKind.SentenceReadText,
                    Snapshot = DocumentState.Snapshot
                };

                DisplaySequenceLocateResult locate = DisplaySequenceLocator.TryResolve(names, locateOptions);
                if (locate == null || locate.MatchCount == 0)
                {
                    return FailResolve(
                        $"locate_failed: 无法定位句子（{locate?.ErrorCode ?? "unknown"}）");
                }

                if (locate.MatchCount > 1 || !locate.DisplayStartPosition.HasValue)
                {
                    return FailResolve(
                        $"locate_failed: 句子编码歧义（{DisplaySequenceLocator.ErrorAmbiguousSentenceCode}），请扩充 codes 或加 in_table");
                }

                int displayStart = locate.DisplayStartPosition.Value;
                Word.Range range = DisplayPositionRangeResolver.ResolveSpan(
                    doc,
                    names,
                    displayStart,
                    target.InTableId,
                    tableScope,
                    locate.OrderedDomainCodes,
                    debugTag: "insert_svg_target");

                if (range == null)
                {
                    return FailResolve("locate_failed: 已解析编码但无法取得 Word Range");
                }

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                return OkRange(range, target, usedCursor: false);
            }
            catch (Exception ex)
            {
                return FailResolve($"locate_failed: {ex.Message}");
            }
        }

        public static Dictionary<string, object> TryGetTargetDict(Dictionary<string, object> args)
        {
            if (args == null || !args.ContainsKey("target") || args["target"] == null)
            {
                return null;
            }

            object raw = args["target"];
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
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

            return echo;
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
            return new ResolveRangeResult { Success = false, Error = error };
        }
    }
}
