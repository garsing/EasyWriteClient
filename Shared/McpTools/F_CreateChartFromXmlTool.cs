#pragma warning disable CS0119

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从 XML 配置文件创建图表。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_CreateChartFromXmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_create_chart_from_xml"] = async (args) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] 开始创建图表");

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_create_chart_from_xml] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    // WPS 对 InlineShapes.AddChart（内嵌 Excel）不稳定，真机 RPC_E_SERVERFAULT；明确失败勿假成功
                    if (docHandle.Context != null && docHandle.Context.Kind == ChannelKind.Wps)
                    {
                        return DocumentHostAdapter.UnsupportedResult(
                            ChannelKind.Wps,
                            docHandle.ChannelId,
                            "create_chart (WPS 不支持 Word/Excel 内嵌图表 COM；请用 Word 或改用 SVG/图片)");
                    }

                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath xmlResolved, out string pathError, "path", "filename"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    string filename = xmlResolved.Display;
                    var xmlRead = await FilePathResolver.ReadAsync(xmlResolved).ConfigureAwait(false);
                    if (!xmlRead.Success)
                    {
                        return new ToolResult { Success = false, Error = xmlRead.Error };
                    }

                    string xmlFilePath = xmlResolved.LocalPath;

                    var parseResult = await ChartXmlParser.ParseXmlFile(xmlFilePath);
                    if (parseResult.Config == null)
                    {
                        return new ToolResult { Success = false, Error = $"XML解析失败: {parseResult.Error}" };
                    }

                    var validationResult = ChartCreateHelper.ValidateConfig(parseResult.Config);
                    if (!validationResult.IsValid)
                    {
                        return new ToolResult { Success = false, Error = $"配置验证失败: {validationResult.Error}" };
                    }

                    int? position = null;
                    if (args.ContainsKey("position"))
                    {
                        if (int.TryParse(args["position"]?.ToString(), out int pos))
                        {
                            position = pos;
                        }
                    }

                    var createResult = ChartCreateHelper.CreateAtLocation(document, parseResult.Config, position);
                    if (!createResult.Success)
                    {
                        return new ToolResult { Success = false, Error = $"图表创建失败: {createResult.Error}" };
                    }

                    await Task.CompletedTask;

                    var warnings = new List<string>();
                    if (!string.IsNullOrEmpty(createResult.Error))
                    {
                        warnings.Add(createResult.Error);
                    }

                    string resolvedChartId = null;
                    try
                    {
                        int chartStart = createResult.ChartStart;
                        WordReader.ReadWord(document);

                        List<Word.InlineShape> allCharts = ChartResolveHelper.GetInlineChartsInOrder(document);
                        int chartIndex = allCharts.FindIndex(s => s.Range.Start == chartStart);

                        var idOrder = DocumentState.GetChartIdOrder();
                        if (chartIndex >= 0 && chartIndex < idOrder.Count)
                        {
                            resolvedChartId = idOrder[chartIndex];
                            System.Diagnostics.Debug.WriteLine(
                                $"[DEBUG] 建图 chart_id={resolvedChartId}（Range.Start={chartStart}, index={chartIndex}）");
                        }
                        else
                        {
                            warnings.Add("未能解析 chart_id（图表索引与编号映射不一致）");
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"未能解析 chart_id: {ex.Message}");
                    }

                    string message = "图表创建成功";
                    if (warnings.Count > 0)
                    {
                        message += $"，但存在以下警告: {string.Join("; ", warnings)}";
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            chart_created = true,
                            chart_id = resolvedChartId,
                            chart_type = parseResult.Config.Chart.Type,
                            title = parseResult.Config.Chart.Title,
                            xml_file = filename,
                            message = message
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }
    }
}
