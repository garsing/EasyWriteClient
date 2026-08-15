using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace WordAddIn1
{
    /// <summary>
    /// WS client_method → 现有 F_* COM handler 映射（复用 McpTools 注册逻辑）。
    /// </summary>
    public class WsMethodRegistry
    {
        private readonly Dictionary<string, string> _methodToToolName;
        private readonly Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> _toolHandlers;
        private readonly JavaScriptSerializer _jsonSerializer = new JavaScriptSerializer();

        public WsMethodRegistry(object wordApplication)
        {
            DocumentCheckpointService.SetWordApplication(wordApplication);

            _toolHandlers = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>();
            McpTools.RegisterBuiltInTools(_toolHandlers, wordApplication);

            _methodToToolName = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "document.getContent", "F_get_document_content" },
                { "document.open", "F_open_document" },
                { "workbook.getContent", "F_get_workbook_content" },
                { "workbook.readRange", "F_read_excel_range" },
                { "workbook.writeRange", "F_write_excel_range" },
                { "workbook.manageSheet", "F_manage_excel_sheet" },
                { "workbook.applyFormat", "F_apply_excel_format" },
                { "workbook.getFormat", "F_get_excel_format" },
                { "workbook.applyConditionalFormat", "F_apply_excel_conditional_format" },
                { "workbook.pivot", "F_excel_pivot" },
                { "channel.setDefault", "F_switch_default_channel" },
                { "document.processActions", "F_process_document_actions" },
                // document.setTrackRevisions / F_set_track_revisions 已下线（2026-08）
                { "document.applyDataProvenance", "F_apply_data_provenance" },
                { "document.getDataProvenance", "F_get_data_provenance" },
                { "document.insertFromFile", "F_insert_document_from_file" },
                { "document.replaceFromFile", "F_replace_document_from_file" },
                { "document.getChunkDisplayContent", "F_get_curr_doc_chunk_display_content" },
                { "document.getChunkParagraphDisplayContent", "F_get_curr_doc_chunk_paragraph_display_content" },
                // { "document.findPosition", "F_find_position" }, // 暂不下线 WS 映射（代码保留）
                { "document.formatTransfer", "F_format_transfer" },
                { "document.applyPageSetup", "F_apply_page_setup" },
                { "document.captureDocumentPageImage", "F_capture_document_page_image" },
                { "document.getPageInfo", "F_get_document_page_info" },
                { "document.listOperations", "F_list_document_operations" },
                { "document.restoreCheckpoint", "F_restore_document_checkpoint" },
                { "table.createFromXml", "F_create_table_from_xml" },
                { "table.applyFormatFromXml", "F_apply_table_format_from_xml" },
                { "table.readRowValues", "F_read_table_row_values" },
                { "table.applyRowValues", "F_apply_table_row_values" },
                { "table.extractFormat", "F_extract_table_format" },
                { "table.optimizeDisplay", "F_optimize_table_display" },
                { "table.delete", "F_delete_table" },
                { "table.applyKbFormat", "F_apply_kb_table_format" },
                { "chart.createFromXml", "F_create_chart_from_xml" },
                { "chart.extractXml", "F_extract_chart_xml" },
                { "chart.applyFromXml", "F_apply_chart_from_xml" },
                { "chart.delete", "F_delete_chart" },
                { "document.insertSvgImage", "F_insert_svg_image" },
                { "document.insertWorkspaceImage", "F_insert_workspace_image" },
                { "file.writeFormat", "F_write_format_file" },
                { "file.read", "F_read_file" },
                { "file.modifyYaml", "F_modify_yaml_file" },
                { "file.csvToXml", "F_csv_to_xml" },
                { "misc.greet", "F_greet" },
                { "document.getFormatContext", "F_get_format_context" },
                { "document.applyDocumentFormat", "F_apply_document_format" },
                { "document.applyParagraphFormat", "F_apply_paragraph_format" },
                { "document.applyKbFormat", "F_apply_kb_format" },
                { "document.applyKbParagraphFormat", "F_apply_kb_paragraph_format" },
                // { "test.wordDocumentExtractor", "F_test_word_document_extractor" },
                { "test.paragraphCodeLocator", "F_test_paragraph_code_locator" },
            };
        }

        public IList<string> GetMethodNames()
        {
            return _methodToToolName.Keys.ToList();
        }

        public async Task<ToolResult> InvokeRawAsync(string method, Dictionary<string, object> parameters)
        {
            if (string.IsNullOrWhiteSpace(method))
            {
                return new ToolResult { Success = false, Error = "method is required" };
            }

            if (!_methodToToolName.TryGetValue(method, out var toolName))
            {
                return new ToolResult { Success = false, Error = $"unknown method '{method}'" };
            }

            if (!_toolHandlers.TryGetValue(toolName, out var handler))
            {
                return new ToolResult { Success = false, Error = $"handler not registered for method '{method}'" };
            }

            // Desktop 缩小版：白名单工具在执行前触发（Plugin 未注册则为 no-op）
            if (CompactLayoutTriggers.ShouldCompact(toolName))
            {
                HostCallbacks.RaiseRequestCompact();
            }

            try
            {
                AgentRunCancellation.ThrowIfCancelled();
            }
            catch (OperationCanceledException)
            {
                return new ToolResult { Success = false, Error = "cancelled by user" };
            }

            var args = parameters ?? new Dictionary<string, object>();

            if (toolName == "F_list_document_operations" || toolName == "F_restore_document_checkpoint")
            {
                return await handler(args);
            }

            if (!DocumentCheckpointMutatingTools.ShouldCheckpoint(toolName, args))
            {
                return await handler(args);
            }

            CheckpointBeforeResult before = DocumentCheckpointService.CreateBeforeMutation(toolName, args);
            if (!before.Success)
            {
                return new ToolResult { Success = false, Error = before.Error };
            }

            ToolResult result;
            try
            {
                result = await handler(args);
            }
            catch (Exception)
            {
                DocumentCheckpointService.RollbackFailedMutation(before.DocUuid, before.Timestamp);
                throw;
            }

            if (DocumentCheckpointService.IsInvokeMutationFailed(result, toolName))
            {
                if (DocumentCheckpointService.IsCheckpointCleanupOnly(result))
                {
                    DocumentCheckpointService.CleanupFailedInvoke(before.DocUuid, before.Timestamp);
                }
                else
                {
                    DocumentCheckpointService.RollbackFailedMutation(before.DocUuid, before.Timestamp);
                }

                return result;
            }

            DocumentCheckpointService.AppendActionLogAfterMutation(toolName, args);
            return result;
        }

        public async Task<string> InvokeAsync(string method, Dictionary<string, object> parameters)
        {
            try
            {
                var result = await InvokeRawAsync(method, parameters);
                return SerializeResult(result);
            }
            catch (Exception ex)
            {
                return SerializeError(ex.Message);
            }
        }

        private string SerializeResult(ToolResult result)
        {
            var payload = new
            {
                success = result.Success,
                message = result.Success ? "工具执行成功" : result.Error,
                data = result.Data
            };
            return _jsonSerializer.Serialize(payload);
        }

        private string SerializeError(string message)
        {
            var payload = new
            {
                success = false,
                message = message,
                data = (object)null
            };
            return _jsonSerializer.Serialize(payload);
        }
    }
}
