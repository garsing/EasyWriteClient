using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// XML 格式迁移：套用 TableConfig XML，仅迁移 General 部分（三字体 + Style）。
    /// </summary>
    public static class F_ApplyTableFormatFromXmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_table_format_from_xml"] = async (args) =>
            {
                try
                {
                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document document, out ToolResult resolveError))
                    {
                        return resolveError;
                    }
                    string tableId = args.ContainsKey("table_id") ? args["table_id"]?.ToString() : "";
                    if (string.IsNullOrEmpty(tableId))
                    {
                        return new ToolResult { Success = false, Error = "未提供表格编号" };
                    }

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

                    var applyResult = TableConfigApplyHelper.ApplyTableConfigToExistingTable(
                        document, tableId, xmlContent);
                    if (!applyResult.Success)
                    {
                        return new ToolResult { Success = false, Error = applyResult.Error };
                    }

                    Word.Application app = wordApplication as Word.Application;
                    Word.Table table = TableConfigApplyHelper.ResolveTableById(document, tableId, out _);
                    if (app != null && table != null)
                    {
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, table.Range, $"apply_table_format:{tableId}");
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_format_applied = true,
                            table_id = tableId,
                            rows = applyResult.RowCount,
                            columns = applyResult.ColumnCount,
                            warnings = applyResult.Warnings,
                            message = "表格格式 apply 成功（仅迁移 General 部分，未改 Table/单元格文字）",
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] F_apply_table_format_from_xml 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }
    }
}
