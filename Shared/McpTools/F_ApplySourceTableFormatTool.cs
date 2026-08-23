using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_ApplySourceTableFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_source_table_format"] = async (args) =>
            {
                try
                {
                    string targetTableId = FormatSourceResolver.GetArgString(args, "target_table_id");
                    string kbTableId = FormatSourceResolver.GetArgString(args, "kb_table_id");

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    Word.Application app = null;
                    try
                    {
                        app = doc.Application;
                    }
                    catch (Exception)
                    {
                    }

                    if (app == null)
                    {
                        app = wordApplication as Word.Application;
                    }

                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取文档 Application（Word/WPS）" };
                    }

                    if (!FormatSourceResolver.TryResolve(
                            args,
                            wordApplication,
                            docHandle.Context?.DocUuid,
                            out FormatSourceResolution source,
                            out ToolResult sourceError))
                    {
                        return sourceError;
                    }

                    return await KbTableFormatApplyHelper.RunApplyAsync(
                        app,
                        targetTableId,
                        kbTableId,
                        source,
                        doc);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_source_table_format 失败: {ex.Message}" };
                }
            };
        }
    }
}
