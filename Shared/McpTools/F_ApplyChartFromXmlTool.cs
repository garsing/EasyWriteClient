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

                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

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
                    string scope = "all";
                    if (args.ContainsKey("scope") && args["scope"] != null)
                    {
                        scope = args["scope"]?.ToString()?.Trim().ToLowerInvariant() ?? "all";
                    }

                    string mode = "full";
                    if (args.ContainsKey("mode") && args["mode"] != null)
                    {
                        mode = args["mode"]?.ToString()?.Trim().ToLowerInvariant() ?? "full";
                    }

                    TryGetXmlContent(args, out string xmlContent, out _);

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
                            Error = $"找不到图表编号 '{chartId}'，请确认图表编号是否正确"
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

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            string chartId = args != null && args.ContainsKey("chart_id")
                ? args["chart_id"]?.ToString()
                : "";
            if (string.IsNullOrEmpty(chartId))
            {
                failure = DocumentCheckpointMutatingTools.Fail("未提供图表编号");
                return true;
            }

            string scope = "all";
            if (args != null && args.ContainsKey("scope") && args["scope"] != null)
            {
                scope = args["scope"]?.ToString()?.Trim().ToLowerInvariant() ?? "all";
            }

            if (scope != "all" && scope != "format" && scope != "data")
            {
                failure = DocumentCheckpointMutatingTools.Fail(
                    "scope 无效: '" + scope + "'，有效值为: all, format, data");
                return true;
            }

            if (!TryGetXmlContent(args, out string xmlContent, out string xmlError))
            {
                failure = DocumentCheckpointMutatingTools.Fail(xmlError);
                return true;
            }

            var parseResult = ChartConfigXmlParser.ParseAsync(xmlContent).GetAwaiter().GetResult();
            if (!parseResult.Success)
            {
                failure = DocumentCheckpointMutatingTools.Fail(parseResult.Error ?? "XML解析失败");
                return true;
            }

            return false;
        }

        private static bool TryGetXmlContent(
            Dictionary<string, object> args,
            out string xmlContent,
            out string error)
        {
            xmlContent = null;
            error = null;
            if (args != null && args.ContainsKey("xml_content") && !string.IsNullOrEmpty(args["xml_content"]?.ToString()))
            {
                xmlContent = args["xml_content"].ToString();
                return true;
            }

            if (FilePathResolver.TryGetArg(args, "path", "filename") != null)
            {
                if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath xmlResolved, out string pathError, "path", "filename"))
                {
                    error = pathError;
                    return false;
                }

                if (!DocumentCheckpointMutatingTools.TryReadLocalText(xmlResolved, out xmlContent, out string readError))
                {
                    error = readError;
                    return false;
                }

                return true;
            }

            error = "必须提供 path 或 xml_content 参数";
            return false;
        }
    }
}
