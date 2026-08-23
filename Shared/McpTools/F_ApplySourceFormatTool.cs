using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_ApplySourceFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_source_format"] = async (args) =>
            {
                try
                {
                    string targetCodes = FormatSourceResolver.GetArgString(args, "target_codes");
                    string kbDetailedSubtype = FormatSourceResolver.GetArgString(args, "kb_detailed_subtype");
                    string tableId = FormatSourceResolver.GetArgString(args, "in_table");
                    if (string.IsNullOrEmpty(tableId))
                    {
                        tableId = null;
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    Word.Application app = ResolveApplication(wordApplication, doc);
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

                    return await KbFormatApplyHelper.RunApplyAsync(
                        app,
                        targetCodes,
                        kbDetailedSubtype,
                        source,
                        tableId,
                        doc);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_source_format 失败: {ex.Message}" };
                }
            };
        }

        private static Word.Application ResolveApplication(object wordApplication, Word.Document doc)
        {
            try
            {
                Word.Application fromDoc = doc?.Application;
                if (fromDoc != null)
                {
                    return fromDoc;
                }
            }
            catch (Exception)
            {
            }

            return wordApplication as Word.Application;
        }
    }
}
