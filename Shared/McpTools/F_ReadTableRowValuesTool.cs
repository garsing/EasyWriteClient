using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 读取表格整表行值（只读）。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ReadTableRowValuesTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_read_table_row_values"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult channelResolveError,
                            activateDocument: false))
                    {
                        return channelResolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_read_table_row_values] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string tableId = args.ContainsKey("table_id") ? args["table_id"]?.ToString() : "";
                    if (string.IsNullOrWhiteSpace(tableId))
                    {
                        return new ToolResult { Success = false, Error = "未提供 table_id" };
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

                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[ReadTableRowValues] 开始读取 table_id={tableId}");
                    TableRowValuesReadResult readResult = TableRowValuesHelper.ReadAllRows(table);
                    if (!readResult.Success)
                    {
                        EasyWriteDiagnostics.LogTableRowValues(
                            $"[ReadTableRowValues] 读取失败: {readResult.Error}");
                        return new ToolResult { Success = false, Error = readResult.Error };
                    }

                    TableRowValuesHelper.LogReadSummary(tableId, readResult, verboseOnly: true);

                    await Task.CompletedTask;

                    string message =
                        $"已读取 {readResult.RowCount} 行；column_count={readResult.ColumnCount}。"
                        + "填空布局 ONLY 用 rows[].values + rows[].slots；"
                        + "禁止参考 getContent 的 <table><cell> HTML。";

                    const string layoutWarning =
                        "表格填空：勿看 getContent 的 <table>/<cell> HTML；"
                        + "本返回的 rows[].values 与 rows[].slots 是 apply 的唯一布局依据。";

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_id = tableId,
                            row_count = readResult.RowCount,
                            column_count = readResult.ColumnCount,
                            rows = readResult.Rows,
                            message,
                            layout_warning = layoutWarning,
                            do_not_use_getcontent_table_html = true,
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ReadTableRowValues] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }
    }
}
