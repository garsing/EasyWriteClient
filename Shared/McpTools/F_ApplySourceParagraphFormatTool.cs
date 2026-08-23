using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_ApplySourceParagraphFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_source_paragraph_format"] = async (args) =>
            {
                try
                {
                    string targetParagraphCodes = FormatSourceResolver.GetArgString(args, "target_paragraph_codes");
                    string clusterId = FormatSourceResolver.GetArgString(args, "cluster_id");
                    string tableId = FormatSourceResolver.GetArgString(args, "in_table");
                    if (string.IsNullOrEmpty(tableId))
                    {
                        tableId = null;
                    }

                    Dictionary<string, object> explicitParaFormat = ParseNestedObject(args, "para_format");

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

                    FormatSourceResolution source = null;
                    if (!string.IsNullOrEmpty(clusterId))
                    {
                        if (!FormatSourceResolver.TryResolve(
                                args,
                                wordApplication,
                                docHandle.Context?.DocUuid,
                                out source,
                                out ToolResult sourceError))
                        {
                            return sourceError;
                        }
                    }

                    return await KbParagraphFormatApplyHelper.RunApplyAsync(
                        app,
                        targetParagraphCodes,
                        clusterId,
                        explicitParaFormat,
                        source,
                        tableId,
                        doc);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_source_paragraph_format 失败: {ex.Message}" };
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

        private static Dictionary<string, object> ParseNestedObject(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return new Dictionary<string, object>();
            }

            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            try
            {
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                return serializer.DeserializeObject(raw.ToString()) as Dictionary<string, object>
                    ?? new Dictionary<string, object>();
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }
    }
}
