using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表格精迁：HTTP 拉 KB 单表 TableConfig XML，对 target_table_id 仅迁移 General 部分。
    /// </summary>
    public static class KbTableFormatApplyHelper
    {
        public static async Task<ToolResult> RunApplyAsync(
            Word.Application wordApp,
            string targetTableId,
            string kbTableId,
            string targetDocumentName,
            string targetKnowledgeBaseUuid,
            string targetStorageDocUuid)
        {
            if (wordApp == null)
            {
                return new ToolResult { Success = false, Error = "Word 应用程序不可用" };
            }

            string tid = (targetTableId ?? "").Trim();
            if (string.IsNullOrEmpty(tid))
            {
                return new ToolResult { Success = false, Error = "缺少 target_table_id" };
            }

            string kbTid = (kbTableId ?? "").Trim();
            if (string.IsNullOrEmpty(kbTid))
            {
                return new ToolResult { Success = false, Error = "缺少 kb_table_id" };
            }

            var storageUuid = (targetStorageDocUuid ?? "").Trim();
            var docName = (targetDocumentName ?? "").Trim();
            var kbUuid = (targetKnowledgeBaseUuid ?? "").Trim();

            if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(docName))
            {
                return new ToolResult { Success = false, Error = "缺少 target_storage_doc_uuid 或 target_document_name" };
            }

            if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(kbUuid))
            {
                return new ToolResult { Success = false, Error = "按 document_name 定位时必须提供 target_knowledge_base_uuid" };
            }

            if (!UserService.Instance.CheckLoginStatus())
            {
                return new ToolResult { Success = false, Error = "用户未登录，无法调用知识库接口" };
            }

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                doc = null;
            }

            if (doc == null)
            {
                return new ToolResult { Success = false, Error = "没有活动的 Word 文档" };
            }

            DocumentState.BindAndActivate(doc);

            try
            {
                WordReader.ReadWord(doc);
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
            }

            if (DocumentState.GetTableIndex(tid) < 0)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = $"找不到当前稿表格编号 '{tid}'，请先 F_get_document_content 刷新 T_ 快照",
                };
            }

            string xmlContent = await FetchTableFormatXmlAsync(
                docName, kbUuid, storageUuid, kbTid);
            if (string.IsNullOrEmpty(xmlContent))
            {
                return new ToolResult
                {
                    Success = false,
                    Error =
                        $"无法获取 KB 表格 XML（kb_table_id={kbTid}）；"
                        + "请先 B_get_kb_table_format_display 核对 catalog 中的 kb_table_id 并原样复制",
                };
            }

            TableConfigApplyResult applyResult = TableConfigApplyHelper.ApplyTableConfigToExistingTable(
                doc, tid, xmlContent, skipReadWord: true);
            if (!applyResult.Success)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = applyResult.Error ?? "表格格式 apply 失败",
                    Data = BuildResultData(tid, kbTid, storageUuid, docName, kbUuid, applyResult, doc),
                };
            }

            var data = BuildResultData(tid, kbTid, storageUuid, docName, kbUuid, applyResult, doc);
            data["message"] = "KB 表格格式精迁成功（仅迁移 General 部分，未改 Table/单元格文字）";
            if (applyResult.RowCount.HasValue)
            {
                data["rows"] = applyResult.RowCount.Value;
            }

            if (applyResult.ColumnCount.HasValue)
            {
                data["columns"] = applyResult.ColumnCount.Value;
            }

            Word.Table table = TableConfigApplyHelper.ResolveTableById(doc, tid, out _);
            if (table != null)
            {
                PostModifyNavigateHelper.NavigateAfterEnd(
                    wordApp, table.Range, $"apply_kb_table_format:{tid}");
            }

            return new ToolResult { Success = true, Data = data };
        }

        public static async Task<string> FetchTableFormatXmlAsync(
            string documentName,
            string knowledgeBaseUuid,
            string storageDocUuid,
            string kbTableId)
        {
            string baseUrl = (ConfigManager.Config.Api.BaseUrl ?? "").TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                System.Diagnostics.Debug.WriteLine("[KbTableFormatApplyHelper] Api.BaseUrl 为空");
                return null;
            }

            string tid = Uri.EscapeDataString((kbTableId ?? "").Trim());
            string url;
            if (!string.IsNullOrEmpty(storageDocUuid))
            {
                url =
                    $"{baseUrl}/knowledge/document/table_format_xml_by_filename?kb_table_id={tid}&storage_doc_uuid={Uri.EscapeDataString(storageDocUuid)}";
            }
            else
            {
                url =
                    $"{baseUrl}/knowledge/document/table_format_xml_by_filename?kb_table_id={tid}&document_name={Uri.EscapeDataString(documentName)}&knowledge_base_uuid={Uri.EscapeDataString(knowledgeBaseUuid)}";
            }

            BackendApiClient.JsonResult result = await BackendApiClient.GetAuthenticatedAsync(
                    url, retryOnUnauthorized: true, timeout: TimeSpan.FromSeconds(120))
                .ConfigureAwait(false);

            if (!result.Success)
            {
                string detail = TryExtractErrorDetail(result.Body);
                System.Diagnostics.Debug.WriteLine(
                    $"[KbTableFormatApplyHelper] GET table_format_xml 失败 {(int)result.StatusCode}: {detail ?? result.Body}");
                return null;
            }

            try
            {
                var root = JObject.Parse(result.Body);
                if (root["success"]?.Value<bool>() != true)
                {
                    return null;
                }

                JToken xml = root["data"]?["xml_content"];
                if (xml == null || xml.Type == JTokenType.Null)
                {
                    return null;
                }

                return xml.ToString();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[KbTableFormatApplyHelper] 解析 table_format_xml 响应失败: {ex.Message}");
                return null;
            }
        }

        private static Dictionary<string, object> BuildResultData(
            string targetTableId,
            string kbTableId,
            string storageUuid,
            string docName,
            string kbUuid,
            TableConfigApplyResult applyResult,
            Word.Document doc)
        {
            var data = new Dictionary<string, object>
            {
                { "target_table_id", targetTableId },
                { "kb_table_id", kbTableId },
                { "target_storage_doc_uuid", storageUuid },
                { "target_document_name", docName },
                { "target_knowledge_base_uuid", kbUuid },
                { "table_format_applied", applyResult.Success },
            };
            if (applyResult.Warnings != null && applyResult.Warnings.Count > 0)
            {
                data["warnings"] = applyResult.Warnings;
            }

            return data;
        }

        private static string TryExtractErrorDetail(string body)
        {
            if (string.IsNullOrWhiteSpace(body))
            {
                return null;
            }

            try
            {
                var root = JObject.Parse(body);
                JToken detail = root["detail"];
                if (detail != null && detail.Type != JTokenType.Null)
                {
                    return detail.ToString();
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }
    }
}
