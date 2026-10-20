using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 按行槽位写入表格数据。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyTableRowValuesTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_table_row_values"] = async (args) =>
            {
                try
                {
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult channelResolveError))
                    {
                        return channelResolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_table_row_values] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string tableId = args.ContainsKey("table_id") ? args["table_id"]?.ToString() : "";
                    ActionFormatHelper.Spec callFormat = ActionFormatHelper.ParseFormatObject(args, required: true);
                    TryParseWriteSpecs(args, out List<TableRowWriteSpec> specs, out _);

                    using (EasyWriteDiagnostics.Time("table_row_values.apply.total"))
                    {
                    try
                    {
                        using (EasyWriteDiagnostics.Time("table_row_values.apply.process_document_before"))
                        {
                            TableRowValuesHelper.EnsureDocumentMapping(
                                document, "table_row_values.apply_before");
                        }
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
                    }

                    Word.Table table;
                    using (EasyWriteDiagnostics.Time("table_row_values.apply.resolve"))
                    {
                        table = TableConfigApplyHelper.ResolveTableById(document, tableId, out string resolveError);
                        if (table == null)
                        {
                            return new ToolResult { Success = false, Error = resolveError };
                        }
                    }

                    Dictionary<string, object> formatSnapshot;
                    string formatMode;
                    using (EasyWriteDiagnostics.Time("table_row_values.apply.format_snapshot"))
                    {
                        if (!ActionFormatHelper.TryBuildSnapshot(
                                document,
                                callFormat,
                                tableScope: null,
                                out formatSnapshot,
                                out formatMode,
                                out string formatError))
                        {
                            return new ToolResult { Success = false, Error = formatError };
                        }
                    }

                    foreach (TableRowWriteSpec spec in specs)
                    {
                        spec.FormatSnapshot = formatSnapshot;
                    }

                    int rowCount;
                    int colCount;
                    List<List<int>> merge;
                    using (EasyWriteDiagnostics.Time("table_row_values.apply.extract_structure"))
                    {
                        TableFormatExtractCore.ExtractStructure(
                            table,
                            out rowCount,
                            out colCount,
                            out merge);
                    }

                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[ApplyTableRowValues] table_id={tableId}, writes={specs.Count}, " +
                        $"rows×cols={rowCount}×{colCount}, format_mode={formatMode}");
                    TableRowValuesHelper.LogWriteSpecs(specs, verboseOnly: true);

                    using (EasyWriteDiagnostics.Time("table_row_values.apply.preflight"))
                    {
                        if (!TableRowValuesHelper.PreflightWrites(
                                table,
                                specs,
                                rowCount,
                                colCount,
                                merge,
                                out string preflightError))
                        {
                            EasyWriteDiagnostics.LogTableRowValues(
                                $"[ApplyTableRowValues] Preflight 失败: {preflightError}");
                            return new ToolResult { Success = false, Error = preflightError };
                        }
                    }

                    TableRowValuesApplyResult applyResult;
                    using (EasyWriteDiagnostics.Time("table_row_values.apply.write"))
                    {
                        applyResult = TableRowValuesHelper.ApplyWrites(
                            table,
                            specs,
                            colCount,
                            merge);
                    }

                    if (!applyResult.Success)
                    {
                        EasyWriteDiagnostics.LogTableRowValues(
                            $"[ApplyTableRowValues] 写入失败: {applyResult.Error}");
                        return new ToolResult { Success = false, Error = applyResult.Error };
                    }

                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[ApplyTableRowValues] 成功 cells_written={applyResult.CellsWritten}"
                        + (applyResult.Warnings.Count > 0
                            ? $", warnings={applyResult.Warnings.Count}"
                            : ""));

                    try
                    {
                        using (EasyWriteDiagnostics.Time("table_row_values.apply.process_document_after"))
                        {
                            TableRowValuesHelper.EnsureDocumentMapping(
                                document, "table_row_values.apply_after");
                        }
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"写后刷新映射失败：{ex.Message}" };
                    }

                    using (EasyWriteDiagnostics.Time("table_row_values.apply.navigate"))
                    {
                        NavigateAfterRowValuesApply(document, tableId, applyResult);
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_id = tableId,
                            cells_written = applyResult.CellsWritten,
                            warnings = applyResult.Warnings,
                            format_mode = formatMode,
                            message =
                                "行值写入成功（可覆盖；<keep/> 跳过；已套显式字符格式）；已刷新文档映射",
                        },
                    };
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyTableRowValues] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            string tableId = args != null && args.ContainsKey("table_id")
                ? args["table_id"]?.ToString()
                : "";
            if (string.IsNullOrWhiteSpace(tableId))
            {
                failure = DocumentCheckpointMutatingTools.Fail("未提供 table_id");
                return true;
            }

            ActionFormatHelper.Spec callFormat = ActionFormatHelper.ParseFormatObject(args, required: true);
            if (!callFormat.Ok)
            {
                failure = DocumentCheckpointMutatingTools.Fail(
                    callFormat.Error ?? ActionFormatHelper.MissingFormatError);
                return true;
            }

            if (!TryParseWriteSpecs(args, out _, out string parseError))
            {
                failure = DocumentCheckpointMutatingTools.Fail(parseError);
                return true;
            }

            return false;
        }

        private static bool TryParseWriteSpecs(
            Dictionary<string, object> args,
            out List<TableRowWriteSpec> specs,
            out string error)
        {
            specs = new List<TableRowWriteSpec>();
            error = null;

            if (args == null || !args.ContainsKey("writes") || args["writes"] == null)
            {
                error = "writes 不能为空";
                return false;
            }

            object writesRaw = args["writes"];
            if (writesRaw is string)
            {
                error = "writes 必须是数组";
                return false;
            }

            IEnumerable writeItems = writesRaw as IEnumerable;
            if (writeItems == null)
            {
                error = "writes 必须是数组";
                return false;
            }

            int itemIndex = 0;
            foreach (object itemObj in writeItems)
            {
                itemIndex++;
                var item = ActionFormatHelper.CoerceDict(itemObj);
                if (item == null)
                {
                    error = $"writes[{itemIndex}] 格式无效";
                    return false;
                }

                if (!item.ContainsKey("row") || item["row"] == null)
                {
                    error = $"writes[{itemIndex}] 缺少 row";
                    return false;
                }

                if (!TryParseInt(item["row"], out int row))
                {
                    error = $"writes[{itemIndex}] row 无效";
                    return false;
                }

                if (!item.ContainsKey("values") || item["values"] == null)
                {
                    error = $"values_required: writes[{itemIndex}]";
                    return false;
                }

                if (!TryParseStringList(item["values"], out List<string> values, out string valuesError))
                {
                    error = valuesError ?? $"writes[{itemIndex}] values 无效";
                    return false;
                }

                if ((item.ContainsKey("anchor") && !string.IsNullOrWhiteSpace(item["anchor"]?.ToString()))
                    || (item.ContainsKey("index") && item["index"] != null))
                {
                    error = TableRowValuesHelper.BuildAnchorRemovedError(row);
                    return false;
                }

                specs.Add(new TableRowWriteSpec
                {
                    Row = row,
                    Values = values,
                });
            }

            if (specs.Count == 0)
            {
                error = "writes 不能为空";
                return false;
            }

            return true;
        }

        private static bool TryParseStringList(object valuesObj, out List<string> values, out string error)
        {
            values = new List<string>();
            error = null;

            if (!(valuesObj is IEnumerable sequence))
            {
                error = "values 必须是数组";
                return false;
            }

            foreach (object valueObj in sequence)
            {
                values.Add(valueObj?.ToString() ?? "");
            }

            return true;
        }

        private static bool TryParseInt(object valueObj, out int value)
        {
            value = 0;
            if (valueObj == null)
            {
                return false;
            }

            try
            {
                value = Convert.ToInt32(valueObj);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void NavigateAfterRowValuesApply(
            Word.Document document,
            string tableId,
            TableRowValuesApplyResult applyResult)
        {
            Word.Application app = null;
            try
            {
                app = document.Application;
            }
            catch
            {
            }

            if (app == null)
            {
                EasyWriteDiagnostics.LogTableRowValues("[ApplyTableRowValues] navigate 无 Application");
                return;
            }

            string tag = $"apply_table_row_values:{tableId}";
            int lastRow = applyResult == null ? 0 : applyResult.LastWrittenRow;
            int lastCol = applyResult == null ? 0 : applyResult.LastWrittenCol;
            int pos = applyResult == null ? 0 : applyResult.LastWrittenNavPos;
            EasyWriteDiagnostics.LogTableRowValues(
                $"[ApplyTableRowValues] navigate last=({lastRow},{lastCol}) pos={pos}");

            Word.Table tableAfter = TableConfigApplyHelper.ResolveTableById(document, tableId, out _);
            if (tableAfter != null)
            {
                Word.Cell lastCell = TableRowValuesHelper.TryResolveWritableCell(
                    tableAfter, lastRow, lastCol);
                if (lastCell != null)
                {
                    PostModifyNavigateHelper.NavigateInsideCellTextEnd(app, lastCell, tag);
                    return;
                }
            }

            if (pos > 0)
            {
                try
                {
                    int storyEnd = document.Content.End;
                    if (pos > storyEnd)
                    {
                        pos = storyEnd;
                    }

                    if (PostModifyNavigateHelper.NavigateToStoryPos(app, pos, tag))
                    {
                        return;
                    }
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[ApplyTableRowValues] navigate pos 失败: {ex.Message}");
                }
            }

            if (tableAfter != null)
            {
                PostModifyNavigateHelper.NavigateAfterEnd(app, tableAfter.Range, tag);
            }
        }
    }
}
