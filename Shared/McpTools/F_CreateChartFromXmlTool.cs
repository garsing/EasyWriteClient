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
                        $"[F_create_chart_from_xml] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    // WPS 对 InlineShapes.AddChart（内嵌 Excel）不稳定，真机 RPC_E_SERVERFAULT；明确失败勿假成功
                    if (docHandle.Context != null && docHandle.Context.Kind == ChannelKind.Wps)
                    {
                        return DocumentHostAdapter.UnsupportedResult(
                            ChannelKind.Wps,
                            docHandle.ChannelId,
                            "create_chart (WPS 不支持 Word/Excel 内嵌图表 COM；请用 Word 或改用 SVG/图片)");
                    }

                    FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath xmlResolved, out _, "path", "filename");
                    string filename = xmlResolved.Display;
                    var parseResult = ChartXmlParser.ParseXmlFile(xmlResolved.LocalPath).GetAwaiter().GetResult();

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

                    Word.InlineShape createdShape = createResult.InlineShape;
                    if (createdShape == null)
                    {
                        return new ToolResult { Success = false, Error = "图表创建失败: 未拿到新图" };
                    }

                    var occupied = new HashSet<string>(StringComparer.Ordinal);
                    foreach (ChartResolveHelper.BodyChart existing in ChartResolveHelper.GetBodyChartsInOrder(document))
                    {
                        if (existing.TryReadTitle(out string title)
                            && ChartTitleStampHelper.TryParse(title, out string existingId, out _))
                        {
                            occupied.Add(existingId);
                        }
                    }

                    string resolvedChartId = ChartTitleStampHelper.NewRandomChartId(occupied);
                    if (string.IsNullOrEmpty(resolvedChartId)
                        || !ChartTitleStampHelper.TryWriteStamp(
                            new ChartResolveHelper.BodyChart { Inline = createdShape, Chart = createdShape.Chart },
                            resolvedChartId))
                    {
                        try
                        {
                            createdShape.Delete();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("[CreateChart] 打戳失败后删图失败: " + ex.Message);
                        }

                        return new ToolResult { Success = false, Error = "图表编号写入形状标题失败" };
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = "生成图表映射表失败：" + ex.Message };
                    }

                    if (DocumentState.GetChartIndex(resolvedChartId) < 0)
                    {
                        return new ToolResult { Success = false, Error = "图表编号未进入正文队列" };
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

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath xmlResolved, out string pathError, "path", "filename"))
            {
                failure = DocumentCheckpointMutatingTools.Fail(pathError);
                return true;
            }

            if (!DocumentCheckpointMutatingTools.TryReadLocalText(xmlResolved, out _, out string readError))
            {
                failure = DocumentCheckpointMutatingTools.Fail(readError);
                return true;
            }

            ChartXmlParseResult parseResult = ChartXmlParser.ParseXmlFile(xmlResolved.LocalPath).GetAwaiter().GetResult();
            if (parseResult.Config == null)
            {
                failure = DocumentCheckpointMutatingTools.Fail("XML解析失败: " + parseResult.Error);
                return true;
            }

            var validationResult = ChartCreateHelper.ValidateConfig(parseResult.Config);
            if (!validationResult.IsValid)
            {
                failure = DocumentCheckpointMutatingTools.Fail("配置验证失败: " + validationResult.Error);
                return true;
            }

            return false;
        }
    }
}
