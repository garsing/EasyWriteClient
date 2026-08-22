using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从文档提取表格 Data（Properties + Data），不含 General（当前未 Register）。
    /// 经 <see cref="DocumentHostAdapter"/>；只读不抢前台。
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
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError,
                            activateDocument: false))
                    {
                        return resolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_extract_table_data] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
                    }

                    var tableSelection = args["table_selection"] as Dictionary<string, object>;
                    var selection = TableExtractSelectionHelper.Resolve(document, tableSelection);
                    if (!selection.Success)
                    {
                        return new ToolResult { Success = false, Error = selection.Error };
                    }

                    TableExtractDto dto;
                    try
                    {
                        dto = TableFormatExtractCore.Extract(selection.Table, document);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExtractTableData] 提取失败：{ex.Message}");
                        return new ToolResult { Success = false, Error = $"表格数据提取失败: {ex.Message}" };
                    }

                    string xmlContent = TableFormatXmlBuilder.BuildDataOnly(dto);
                    int emptyCellCount = TableMergeHelper.CountFillableEmptyCells(dto.Data, dto.Merge);
                    string filename = FilePathResolver.TryGetArg(args, "path")
                        ?? TableFormatFileHelper.GenerateFilename(
                        document,
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
