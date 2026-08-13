using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表格精迁：套用 KB TableConfig XML，仅迁移 General 部分（三字体 + Style）。
    /// 经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyKbTableFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_kb_table_format"] = async (args) =>
            {
                try
                {
                    string targetTableId = args.TryGetValue("target_table_id", out object tgtObj) && tgtObj != null
                        ? tgtObj.ToString()?.Trim() ?? ""
                        : "";
                    string kbTableId = args.TryGetValue("kb_table_id", out object kbObj) && kbObj != null
                        ? kbObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetStorageDocUuid = args.TryGetValue("target_storage_doc_uuid", out object sidObj) && sidObj != null
                        ? sidObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetDocumentName = args.TryGetValue("target_document_name", out object nameObj) && nameObj != null
                        ? nameObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetKnowledgeBaseUuid = args.TryGetValue("target_knowledge_base_uuid", out object kbUuidObj) && kbUuidObj != null
                        ? kbUuidObj.ToString()?.Trim() ?? ""
                        : "";

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_apply_kb_table_format] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_kb_table_format] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

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

                    return await KbTableFormatApplyHelper.RunApplyAsync(
                        app,
                        targetTableId,
                        kbTableId,
                        targetDocumentName,
                        targetKnowledgeBaseUuid,
                        targetStorageDocUuid,
                        doc);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_kb_table_format 失败: {ex.Message}" };
                }
            };
        }
    }
}
