using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 按 table_id 删除整张表格；删后光标落原表后沿，供 F_create_table_from_xml 重建。
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
                    if (string.IsNullOrWhiteSpace(tableId))
                    {
                        return new ToolResult { Success = false, Error = "未提供表格编号" };
                    }
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

                    int rows = targetTable.Rows.Count;
                    int columns = targetTable.Columns.Count;
                    int endPos = targetTable.Range.End;
                    bool hadNestedTables = TableHasNestedTables(targetTable);

                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteTable] 即将删除 table_id={tableId}, endPos={endPos}, " +
                        $"rows×cols={rows}×{columns}, had_nested_tables={hadNestedTables}");

                    try
                    {
                        targetTable.Delete();
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"删除表格失败：{ex.Message}" };
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"删表后刷新映射失败：{ex.Message}" };
                    }

                    Word.Application app = wordApplication as Word.Application;
                    if (app != null)
                    {
                        int docEnd = document.Content.End;
                        int pos = endPos < 0 ? 0 : (endPos > docEnd ? docEnd : endPos);
                        Word.Range afterTableRange = document.Range(pos, pos);
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, afterTableRange, $"delete_table:{tableId}");
                    }

                    string message = BuildSuccessMessage(tableId, rows, columns, hadNestedTables);

                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteTable] 成功 table_id={tableId}, 光标落点 endPos={endPos}");

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_deleted = true,
                            table_id = tableId,
                            rows = rows,
                            columns = columns,
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

        private static string BuildSuccessMessage(string tableId, int rows, int columns, bool hadNestedTables)
        {
            var sb = new StringBuilder();
            sb.Append($"表格 {tableId} 已删除（原 {rows} 行 × {columns} 列）。");
            sb.Append("光标已停留在被删表格原后沿处。");
            sb.Append("请在此位置调用 F_create_table_from_xml 重建新表（无需传入 position 参数）。");
            if (hadNestedTables)
            {
                sb.Append("（该表内曾含嵌套子表，已随本表一并删除。）");
            }

            return sb.ToString();
        }
    }
}
