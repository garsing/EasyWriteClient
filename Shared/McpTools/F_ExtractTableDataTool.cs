using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从 Word 文档提取表格 Data（Properties + Data），不含 General。
    /// </summary>
    public static class F_ExtractTableDataTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_extract_table_data"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序实例不可用" };
                    }

                    dynamic wordApp = wordApplication;
                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    DocumentState.BindAndActivate(wordApp.ActiveDocument);

                    try
                    {
                        WordReader.ReadWord(wordApp.ActiveDocument);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
                    }

                    var tableSelection = args["table_selection"] as Dictionary<string, object>;
                    var selection = TableExtractSelectionHelper.Resolve(wordApp.ActiveDocument, tableSelection);
                    if (!selection.Success)
                    {
                        return new ToolResult { Success = false, Error = selection.Error };
                    }

                    TableExtractDto dto;
                    try
                    {
                        dto = TableFormatExtractCore.Extract(selection.Table, wordApp.ActiveDocument);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExtractTableData] 提取失败：{ex.Message}");
                        return new ToolResult { Success = false, Error = $"表格数据提取失败: {ex.Message}" };
                    }

                    string xmlContent = TableFormatXmlBuilder.BuildDataOnly(dto);
                    int emptyCellCount = TableMergeHelper.CountFillableEmptyCells(dto.Data, dto.Merge);
                    string filename = TableFormatFileHelper.GenerateFilename(
                        wordApp.ActiveDocument,
                        selection.TableIndex,
                        "data");
                    var (saved, savedContent) = await TableFormatFileHelper.SaveXmlAsync(filename, xmlContent);

                    return new ToolResult
                    {
                        Success = saved,
                        Error = saved ? null : "表格数据保存失败",
                        Data = new
                        {
                            data_extracted = saved,
                            table_info = new
                            {
                                rows = dto.Rows,
                                columns = dto.Cols,
                                table_index = selection.TableIndex,
                                table_id = selection.TableId,
                                selection_type = selection.SelectionType,
                            },
                            data_file = filename,
                            data_saved = saved,
                            data_content = savedContent ?? xmlContent,
                            empty_cell_count = emptyCellCount,
                            message = saved
                                ? $"表格数据提取并保存成功，文件名为：{filename}。empty_cell_count 不含 merge 延续格。"
                                : "表格数据保存失败",
                        },
                    };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"表格数据提取失败: {ex.Message}" };
                }
            };
        }
    }
}
