using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 按 table_id 删除整张表格；删后光标落原表开头，供 F_create_table_from_xml 重建。
    /// 经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_DeleteTableTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_delete_table"] = async (args) =>
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
                        $"[F_delete_table] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string tableId = args.ContainsKey("table_id") ? args["table_id"]?.ToString() : "";
                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
                    }

                    Word.Table targetTable = TableConfigApplyHelper.ResolveTableById(document, tableId, out string resolveError);
                    if (targetTable == null)
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    int startPos = TryGetRangeStart(targetTable);
                    bool hadNestedTables = TableHasNestedTables(targetTable);

                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteTable] 即将删除 table_id={tableId}, startPos={startPos}, " +
                        $"had_nested_tables={hadNestedTables}");

                    try
                    {
                        targetTable.Delete();
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"删除表格失败：{ex.Message}" };
                    }

                    bool mappingRefreshed = true;
                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        mappingRefreshed = false;
                        System.Diagnostics.Debug.WriteLine(
                            $"[DeleteTable] 删后刷新映射失败（表已删，不回滚）: {ex.Message}");
                    }

                    bool cursorMoved = TryNavigateToStart(
                        wordApplication as Word.Application, document, startPos, tableId);

                    string message = BuildSuccessMessage(
                        tableId, hadNestedTables, cursorMoved, mappingRefreshed);

                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteTable] 成功 table_id={tableId}, 光标落点 startPos={startPos}, " +
                        $"cursor_moved={cursorMoved}, mapping_refreshed={mappingRefreshed}");

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_deleted = true,
                            table_id = tableId,
                            had_nested_tables = hadNestedTables,
                            message = message,
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DeleteTable] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }

        private static int TryGetRangeStart(Word.Table table)
        {
            try
            {
                return table.Range.Start;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DeleteTable] 读取 Range.Start 失败: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// 删后落到原表开头（D7）。Word 合法光标上限是 Content.End - 1。
        /// 落点失败不得让已成功的 Delete 整单失败（否则 checkpoint 会把表灌回来）。
        /// </summary>
        private static bool TryNavigateToStart(
            Word.Application app,
            Word.Document document,
            int startPos,
            string tableId)
        {
            if (app == null || document == null)
            {
                return false;
            }

            try
            {
                int maxPos = Math.Max(0, document.Content.End - 1);
                int pos = startPos < 0 ? maxPos : Math.Min(Math.Max(0, startPos), maxPos);
                Word.Range atStart = document.Range(pos, pos);
                return PostModifyNavigateHelper.NavigateAfterEnd(
                    app, atStart, $"delete_table:{tableId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[DeleteTable] 移光标失败（表已删，不回滚）: {ex.Message}");
                return false;
            }
        }

        private static bool TableHasNestedTables(Word.Table table)
        {
            if (table == null)
            {
                return false;
            }

            try
            {
                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        if (cell.Tables != null && cell.Tables.Count > 0)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DeleteTable] 嵌套检测异常: {ex.Message}");
            }

            return false;
        }

        private static string BuildSuccessMessage(
            string tableId,
            bool hadNestedTables,
            bool cursorMoved,
            bool mappingRefreshed)
        {
            var sb = new StringBuilder();
            sb.Append($"表格 {tableId} 已删除。");

            if (cursorMoved)
            {
                sb.Append("光标已停留在被删表格原开头处。");
                sb.Append("请在此位置调用 F_create_table_from_xml 重建新表（无需传入 target）。");
            }
            else
            {
                sb.Append("未能将光标移到原表开头。请自行定位后调用 F_create_table_from_xml 重建新表（无需传入 target）。");
            }

            if (hadNestedTables)
            {
                sb.Append("（该表内曾含嵌套子表，已随本表一并删除。）");
            }

            if (!mappingRefreshed)
            {
                sb.Append("删后映射刷新失败，建议调用 F_get_document_content。");
            }

            return sb.ToString();
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
                failure = DocumentCheckpointMutatingTools.Fail("未提供表格编号");
                return true;
            }

            return false;
        }
    }
}
