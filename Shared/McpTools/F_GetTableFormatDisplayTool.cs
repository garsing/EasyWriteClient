using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_GetTableFormatDisplayTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_table_format_display"] = async (args) =>
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
                        FormatSourceEnsure.EnsureTableCatalog(sourceDoc, source.ForceRefresh, source.DisallowBackendApi);
                    });

                    DocumentSessionState session = FormatSourceEnsure.SessionOf(sourceDoc);
                    var entries = new List<Dictionary<string, object>>();
                    if (session?.TableFormatCatalog != null)
                    {
                        foreach (Dictionary<string, object> row in session.TableFormatCatalog)
                        {
                            if (row == null)
                            {
                                continue;
                            }

                            var view = new Dictionary<string, object>(row, StringComparer.Ordinal);
                            view.Remove("xml_content");
                            entries.Add(view);
                        }
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["source_channel_id"] = source.ChannelId,
                        ["entries"] = entries,
                        ["table_format_standard"] = session?.TableFormatStandard
                    };

                    return await Task.FromResult(new ToolResult { Success = true, Data = data });
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "F_get_table_format_display 失败: " + ex.Message };
                }
            };
        }
    }
}
