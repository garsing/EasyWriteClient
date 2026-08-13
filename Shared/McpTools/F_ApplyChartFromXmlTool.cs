using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 将 ChartConfig XML 套用到已有图表。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyChartFromXmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_chart_from_xml"] = async (args) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[ApplyChart] 开始套用图表 XML");

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
                        $"[F_apply_chart_from_xml] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    if (docHandle.Context != null && docHandle.Context.Kind == ChannelKind.Wps)
                    {
                        return DocumentHostAdapter.UnsupportedResult(
                            ChannelKind.Wps,
                            docHandle.ChannelId,
                            "apply_chart (WPS 不支持 Word/Excel 内嵌图表 COM；请用 Word)");
                    }

                    string chartId = args.ContainsKey("chart_id") ? args["chart_id"]?.ToString() : "";
                    if (string.IsNullOrEmpty(chartId))
                    {
                        return new ToolResult { Success = false, Error = "未提供图表编号" };
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

                    string mode = "full";
                    if (args.ContainsKey("mode") && args["mode"] != null)
                    {
                        mode = args["mode"]?.ToString()?.Trim().ToLowerInvariant() ?? "full";
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

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成图表映射表失败：{ex.Message}" };
                    }

                    Word.InlineShape inlineShape = ChartResolveHelper.ResolveInlineShapeByChartId(document, chartId);
                    if (inlineShape == null)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = $"无法定位图表 InlineShape（chart_id={chartId}）"
                        };
                    }

                    // 尽早断开外链，避免 Validate / FillChartData 访问 ChartData 时弹出链接不可用对话框
                    Word.Application application = document.Application;
                    Word.WdAlertLevel previousAlerts = application.DisplayAlerts;
                    try
                    {
                        application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        ChartComUiHelper.TryBreakExternalChartLink(inlineShape.Chart);
                    }
                    finally
                    {
                        application.DisplayAlerts = previousAlerts;
                    }

                    var parseResult = await ChartConfigXmlParser.ParseAsync(xmlContent);
                    if (!parseResult.Success)
                    {
                        return new ToolResult { Success = false, Error = parseResult.Error ?? "XML解析失败" };
                    }

                    var validation = ChartStructureValidator.Validate(
                        inlineShape.Chart,
                        parseResult.Config.Chart,
                        scope);
                    if (!validation.Success)
                    {
                        return new ToolResult { Success = false, Error = validation.Error };
                    }

                    var applyResult = ChartApplyHelper.Apply(
                        inlineShape.Chart,
                        inlineShape,
                        parseResult.Config.Chart,
                        scope,
                        mode);
                    if (!applyResult.Success)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = applyResult.Error,
                            Data = new { chart_id = chartId, scope = scope }
                        };
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyChart] 套用后 ReadWord 失败: {ex.Message}");
                    }

                    PostModifyNavigateHelper.NavigateAfterEnd(
                        application, inlineShape.Range, $"apply_chart:{chartId}");

                    await Task.CompletedTask;

                    string message = $"图表套用成功（scope={scope}）";
                    if (applyResult.Warnings != null && applyResult.Warnings.Count > 0)
                    {
                        message += $"，但存在以下警告: {string.Join("; ", applyResult.Warnings)}";
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            chart_applied = true,
                            chart_id = chartId,
                            scope = scope,
                            mode = mode,
                            series_updated = applyResult.SeriesUpdated,
                            points_updated = applyResult.PointsUpdated,
                            data_source = parseResult.DataSource,
                            csv_file = parseResult.CsvFileName,
                            warnings = applyResult.Warnings,
                            message = message
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyChart] 失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"图表套用失败: {ex.Message}" };
                }
            };
        }
    }
}
