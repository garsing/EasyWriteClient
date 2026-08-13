using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从图表提取 ChartConfig XML（只读）。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ExtractChartXmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_extract_chart_xml"] = async (args) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[ExtractChart] 开始提取图表 XML");

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
                        $"[F_extract_chart_xml] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    if (docHandle.Context != null && docHandle.Context.Kind == ChannelKind.Wps)
                    {
                        return DocumentHostAdapter.UnsupportedResult(
                            ChannelKind.Wps,
                            docHandle.ChannelId,
                            "extract_chart (WPS 不支持 Word/Excel 内嵌图表 COM；请用 Word)");
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成图表映射表失败：{ex.Message}" };
                    }

                    var chartSelection = args["chart_selection"] as Dictionary<string, object>;
                    var selection = ChartSelectionHelper.Resolve(document, chartSelection);
                    if (!selection.Success)
                    {
                        return new ToolResult { Success = false, Error = selection.Error };
                    }

                    string scope = "all";
                    if (args.ContainsKey("scope") && args["scope"] != null)
                    {
                        scope = args["scope"]?.ToString()?.Trim().ToLowerInvariant() ?? "all";
                    }

                    if (scope != "all" && scope != "format" && scope != "data")
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = $"scope 无效: '{scope}'，有效值为: all, format, data"
                        };
                    }

                    bool saveToFile = true;
                    if (args.ContainsKey("save_to_file"))
                    {
                        saveToFile = ParseBoolArg(args["save_to_file"], true);
                    }

                    Word.InlineShape inlineShape = ChartResolveHelper.ResolveInlineShapeByChartId(
                        document,
                        selection.ChartId);
                    if (inlineShape == null)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = $"无法定位图表 InlineShape（chart_id={selection.ChartId}）"
                        };
                    }

                    Word.Chart chart = inlineShape.Chart;
                    var config = ChartExtractHelper.Extract(chart, inlineShape, scope);
                    string xmlContent = ChartFormatXmlBuilder.Build(config, scope);

                    string xmlFile = null;
                    bool saved = false;
                    if (saveToFile)
                    {
                        string filename = ChartFormatXmlBuilder.GenerateFilename(scope, selection.ChartId);
                        var saveResult = await TableFormatFileHelper.SaveXmlAsync(filename, xmlContent);
                        saved = saveResult.Success;
                        xmlFile = filename;
                        if (!saved)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = $"图表 XML 保存失败（filename={xmlFile}，请检查 chart_id 是否含路径非法字符）",
                                Data = new
                                {
                                    xml_content = xmlContent,
                                    scope = scope,
                                    chart_id = selection.ChartId,
                                    xml_file = xmlFile
                                }
                            };
                        }
                    }

                    string message = saveToFile
                        ? $"图表 XML（scope={scope}）提取并保存成功，文件名为：{xmlFile}"
                        : $"图表 XML（scope={scope}）提取成功（未落盘）";

                    System.Diagnostics.Debug.WriteLine(
                        $"[ExtractChart] chart_id={selection.ChartId} scope={scope} save={saveToFile}");

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            chart_extracted = true,
                            chart_info = new
                            {
                                chart_index = selection.ChartIndex,
                                chart_id = selection.ChartId,
                                selection_type = selection.SelectionType,
                                scope = scope,
                                chart_type = config?.Chart?.Type
                            },
                            xml_content = xmlContent,
                            xml_file = xmlFile,
                            format_saved = saved,
                            message = message
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExtractChart] 失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"图表 XML 提取失败: {ex.Message}" };
                }
            };
        }

        private static bool ParseBoolArg(object value, bool defaultValue)
        {
            if (value == null)
            {
                return defaultValue;
            }

            if (value is bool b)
            {
                return b;
            }

            string s = value.ToString()?.Trim().ToLowerInvariant();
            if (s == "true" || s == "1")
            {
                return true;
            }

            if (s == "false" || s == "0")
            {
                return false;
            }

            return defaultValue;
        }
    }
}
