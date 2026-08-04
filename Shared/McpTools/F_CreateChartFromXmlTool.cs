#pragma warning disable CS0119

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从 XML 配置文件创建 Word 图表（MCP 工具注册层）。
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

                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document document, out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    string filename = args.ContainsKey("filename") ? args["filename"]?.ToString() : "";
                    if (string.IsNullOrEmpty(filename))
                    {
                        return new ToolResult { Success = false, Error = "未提供文件名" };
                    }

                    var userService = UserService.Instance;
                    if (!userService.CheckLoginStatus()
                        || string.IsNullOrWhiteSpace(userService.WorkspaceRootEffective))
                    {
                        return new ToolResult { Success = false, Error = "用户未登录或工作区未初始化" };
                    }

                    var ensure = await McpToolsHelpers.EnsureWorkspaceFileAsync(filename).ConfigureAwait(false);
                    if (!ensure.success)
                    {
                        return new ToolResult { Success = false, Error = ensure.error };
                    }

                    string xmlFilePath = ensure.localPath;

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
