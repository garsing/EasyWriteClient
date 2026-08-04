using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 对已有表格 apply Table/Data XML（仅 Data 矩阵，不改 General/结构）。
    /// </summary>
    public static class F_ApplyTableDataFromXmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_table_data_from_xml"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序不可用" };
                    }

                    dynamic wordApp = wordApplication;
                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    DocumentState.BindAndActivate(wordApp.ActiveDocument);
                    Word.Document document = wordApp.ActiveDocument;

                    string tableId = args.ContainsKey("table_id") ? args["table_id"]?.ToString() : "";
                    if (string.IsNullOrEmpty(tableId))
                    {
                        return new ToolResult { Success = false, Error = "未提供 table_id" };
                    }

                    string mode = args.ContainsKey("mode") ? args["mode"]?.ToString() : "empty_only";

                    string xmlContent = null;
                    if (args.ContainsKey("xml_content") && !string.IsNullOrEmpty(args["xml_content"]?.ToString()))
                    {
                        xmlContent = args["xml_content"].ToString();
                    }
                    else if (args.ContainsKey("filename") && !string.IsNullOrEmpty(args["filename"]?.ToString()))
                    {
                        string filename = args["filename"].ToString();
                        if (!UserService.Instance.CheckLoginStatus()
                            || string.IsNullOrWhiteSpace(UserService.Instance.WorkspaceRootEffective))
                        {
                            return new ToolResult { Success = false, Error = "用户未登录或工作区未初始化" };
                        }

                        var ensure = await McpToolsHelpers.EnsureWorkspaceFileAsync(filename).ConfigureAwait(false);
                        if (!ensure.success)
                        {
                            return new ToolResult { Success = false, Error = ensure.error };
                        }

                        xmlContent = File.ReadAllText(ensure.localPath, Encoding.UTF8);
                    }
                    else
                    {
                        return new ToolResult { Success = false, Error = "必须提供 filename 或 xml_content 参数" };
                    }

                    var applyResult = TableDataApplyHelper.ApplyTableDataToExistingTable(
                        document,
                        tableId,
                        xmlContent,
                        mode);

                    if (!applyResult.Success)
                    {
                        return new ToolResult { Success = false, Error = applyResult.Error };
                    }

                    Word.Application app = wordApp as Word.Application;
                    Word.Table table = TableConfigApplyHelper.ResolveTableById(document, tableId, out _);
                    if (app != null && table != null)
                    {
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, table.Range, $"apply_table_data:{tableId}");
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            data_applied = true,
                            table_id = tableId,
                            mode = applyResult.Mode,
                            cells_written = applyResult.CellsWritten,
                            cells_skipped = applyResult.CellsSkipped,
                            warnings = applyResult.Warnings,
                            message =
                                "表格 Data apply 成功，已 ReadWord；若需 S_ 编码可再调 F_get_document_content",
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyTableData] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }
    }
}
