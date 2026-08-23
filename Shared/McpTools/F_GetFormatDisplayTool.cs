using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_GetFormatDisplayTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_format_display"] = async (args) =>
            {
                try
                {
                    if (!FormatSourceResolver.TryResolve(
                            args,
                            wordApplication,
                            targetDocUuid: null,
                            out FormatSourceResolution source,
                            out ToolResult sourceError,
                            activateSource: false))
                    {
                        return sourceError;
                    }

                    if (source.Kind != FormatSourceKind.Local)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "C# 仅处理 source_channel_id；知识库路径由服务端门面分流"
                        };
                    }

                    Word.Document sourceDoc = source.Document;
                    FormatSourceEnsure.RunOnSource(sourceDoc, null, () =>
                    {
                        FormatSourceEnsure.EnsureCharMapping(sourceDoc, source.ForceRefresh, source.DisallowBackendApi);
                        FormatSourceEnsure.EnsurePageLayout(sourceDoc, source.ForceRefresh);
                        FormatSourceEnsure.EnsureParaClusters(sourceDoc, source.ForceRefresh, source.DisallowBackendApi);
                    });

                    DocumentSessionState session = FormatSourceEnsure.SessionOf(sourceDoc);
                    var entries = new List<Dictionary<string, object>>();
                    if (session?.SubtypeFormatMapping != null)
                    {
                        foreach (Dictionary<string, object> item in session.SubtypeFormatMapping)
                        {
                            if (item == null)
                            {
                                continue;
                            }

                            string dst = item.TryGetValue("detailed_subtype", out object dstRaw)
                                ? dstRaw?.ToString() ?? ""
                                : "";
                            if (string.IsNullOrEmpty(dst))
                            {
                                continue;
                            }

                            object format = null;
                            if (item.TryGetValue("format", out object fmt))
                            {
                                format = fmt;
                            }

                            entries.Add(new Dictionary<string, object>
                            {
                                ["detailed_subtype"] = dst,
                                ["type"] = item.ContainsKey("type") ? item["type"] : null,
                                ["level"] = item.ContainsKey("level") ? item["level"] : null,
                                ["char_format"] = format,
                                ["display_label"] = dst
                            });
                        }
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["source_channel_id"] = source.ChannelId,
                        ["entries"] = entries,
                        ["page_layout_summary"] = session?.SectionPageLayout,
                        ["para_format_clusters"] = session?.ParaFormatClusters
                            ?? new Dictionary<string, object> { ["missing"] = true }
                    };

                    return await Task.FromResult(new ToolResult { Success = true, Data = data });
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "F_get_format_display 失败: " + ex.Message };
                }
            };
        }
    }
}
