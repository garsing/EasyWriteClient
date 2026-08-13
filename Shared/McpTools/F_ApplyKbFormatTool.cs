using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 精迁方式 1：按 KB detailed_subtype 对 target_codes explicit apply 字符格式。
    /// 文档触点经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyKbFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_kb_format"] = async (args) =>
            {
                try
                {
                    string targetCodes = args.TryGetValue("target_codes", out object codesObj) && codesObj != null
                        ? codesObj.ToString()?.Trim() ?? ""
                        : "";
                    string kbDetailedSubtype = args.TryGetValue("kb_detailed_subtype", out object subtypeObj) && subtypeObj != null
                        ? subtypeObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetStorageDocUuid = args.TryGetValue("target_storage_doc_uuid", out object sidObj) && sidObj != null
                        ? sidObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetDocumentName = args.TryGetValue("target_document_name", out object nameObj) && nameObj != null
                        ? nameObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetKnowledgeBaseUuid = args.TryGetValue("target_knowledge_base_uuid", out object kbObj) && kbObj != null
                        ? kbObj.ToString()?.Trim() ?? ""
                        : "";
                    string tableId = args.TryGetValue("in_table", out object tableObj) && tableObj != null
                        ? tableObj.ToString()?.Trim()
                        : null;

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_apply_kb_format] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_kb_format] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    Word.Application app = ResolveApplication(wordApplication, doc);
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取文档 Application（Word/WPS）" };
                    }

                    return await KbFormatApplyHelper.RunApplyAsync(
                        app,
                        targetCodes,
                        kbDetailedSubtype,
                        targetDocumentName,
                        targetKnowledgeBaseUuid,
                        targetStorageDocUuid,
                        tableId,
                        doc);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_kb_format 失败: {ex.Message}" };
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
