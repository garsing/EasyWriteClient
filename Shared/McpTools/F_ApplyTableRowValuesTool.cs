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
                    if (string.IsNullOrWhiteSpace(tableId))
                    {
                        return new ToolResult { Success = false, Error = "未提供 table_id" };
                    }

                    ActionFormatHelper.Spec callFormat = ActionFormatHelper.ParseFormatObject(args, required: true);
                    if (!callFormat.Ok)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = callFormat.Error ?? "必须指定 format（inherit_from 或 char_format）"
                        };
                    }

                    if (!TryParseWriteSpecs(args, out List<TableRowWriteSpec> specs, out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
                    }

                    Word.Table table = TableConfigApplyHelper.ResolveTableById(document, tableId, out string resolveError);
                    if (table == null)
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!ActionFormatHelper.TryBuildSnapshot(
                            document,
                            callFormat,
                            tableScope: null,
                            out Dictionary<string, object> formatSnapshot,
                            out string formatMode,
                            out string formatError))
                    {
                        return new ToolResult { Success = false, Error = formatError };
                    }

                    foreach (TableRowWriteSpec spec in specs)
                    {
                        spec.FormatSnapshot = formatSnapshot;
                    }

                    TableFormatExtractCore.ExtractStructure(
                        table,
                        out int rowCount,
                        out int colCount,
                        out List<List<int>> merge);

                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[ApplyTableRowValues] table_id={tableId}, writes={specs.Count}, " +
                        $"rows×cols={rowCount}×{colCount}, format_mode={formatMode}");
                    TableRowValuesHelper.LogWriteSpecs(specs, verboseOnly: true);

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

                    TableRowValuesApplyResult applyResult = TableRowValuesHelper.ApplyWrites(
                        table,
                        specs,
                        colCount,
                        merge);
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
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"写后刷新映射失败：{ex.Message}" };
                    }

                    Word.Application app = wordApplication as Word.Application;
                    Word.Table tableAfter = TableConfigApplyHelper.ResolveTableById(document, tableId, out _);
                    if (app != null && tableAfter != null)
                    {
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, tableAfter.Range, $"apply_table_row_values:{tableId}");
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
                                "行值写入成功（仅空槽，已套显式字符格式）；已 ReadWord；改已有内容请用 F_process_document_actions",
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyTableRowValues] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
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

            IEnumerable writeItems = args["writes"] as IEnumerable;
            if (writeItems == null)
            {
                error = "writes 必须是数组";
                return false;
            }

            int itemIndex = 0;
            foreach (object itemObj in writeItems)
            {
                itemIndex++;
                var item = itemObj as Dictionary<string, object>;
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
    }
}
