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
        private readonly object _wordApplication;
        private readonly Dictionary<string, string> _methodToToolName;
        private readonly Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> _toolHandlers;
        private readonly JavaScriptSerializer _jsonSerializer = new JavaScriptSerializer();

        public string ResolveToolName(string method)
        {
            if (string.IsNullOrWhiteSpace(method))
            {
                return method;
            }

            string key = method.Trim();
            string mapped;
            if (_methodToToolName.TryGetValue(key, out mapped) && !string.IsNullOrEmpty(mapped))
            {
                return mapped;
            }

            return key;
        }

        public WsMethodRegistry(object wordApplication)
        {
            _wordApplication = wordApplication;
            DocumentCheckpointService.SetWordApplication(wordApplication);

            _toolHandlers = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>();
            McpTools.RegisterBuiltInTools(_toolHandlers, wordApplication);

            _methodToToolName = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "document.getContent", "F_get_document_content" },
                { "document.open", "F_open_document" },
                { "terminal.run", "F_run_terminal" },
                { "terminal.close", "F_close_terminal" },
                { "document.close", "F_close_document" },
                { "browser.navigate", "F_browser_navigate" },
                { "browser.snapshot", "F_browser_snapshot" },
                { "browser.interact", "F_browser_interact" },
                { "browser.download", "F_browser_download" },
                { "workbook.getContent", "F_get_workbook_content" },
                { "presentation.getContent", "F_get_presentation_content" },
                { "presentation.readHtml", "F_read_ppt_html" },
                { "presentation.applyHtml", "F_apply_ppt_html" },
                { "presentation.manageSlide", "F_manage_ppt_slide" },
                { "presentation.manageShape", "F_manage_ppt_shape" },
                { "presentation.manageLib", "F_manage_ppt_lib" },
                { "presentation.animation", "F_ppt_animation" },
                { "presentation.transition", "F_ppt_transition" },
                { "workbook.readRange", "F_read_excel_range" },
                { "workbook.writeRange", "F_write_excel_range" },
                { "workbook.manageSheet", "F_manage_excel_sheet" },
                { "workbook.applyFormat", "F_apply_excel_format" },
                { "workbook.getFormat", "F_get_excel_format" },
                { "workbook.applyConditionalFormat", "F_apply_excel_conditional_format" },
                { "workbook.applyStructure", "F_apply_excel_structure" },
                { "workbook.pivot", "F_excel_pivot" },
                { "workbook.chart", "F_excel_chart" },
                { "channel.setDefault", "F_switch_default_channel" },
                { "document.processActions", "F_process_document_actions" },
                // document.setTrackRevisions / F_set_track_revisions 已下线（2026-08）
                { "document.applyDataProvenance", "F_apply_data_provenance" },
                { "document.clearDataProvenance", "F_clear_data_provenance" },
                { "document.getDataProvenance", "F_get_data_provenance" },
                { "document.insertFromFile", "F_insert_document_from_file" },
                { "document.replaceFromFile", "F_replace_document_from_file" },
                // document.getChunkDisplayContent / F_get_curr_doc_chunk_* 已下线（2026-10）：实现保留
                // { "document.getChunkDisplayContent", "F_get_curr_doc_chunk_display_content" },
                // { "document.getChunkParagraphDisplayContent", "F_get_curr_doc_chunk_paragraph_display_content" },
                // { "document.findPosition", "F_find_position" }, // 暂不下线 WS 映射（代码保留）
                { "document.formatTransfer", "F_format_transfer" },
                { "document.applyPageSetup", "F_apply_page_setup" },
                { "document.captureDocumentPageImage", "F_capture_document_page_image" },
                { "presentation.capturePptShape", "F_capture_ppt_shape" },
                { "document.getPageInfo", "F_get_document_page_info" },
                { "document.insertBlankPage", "F_insert_blank_page" },
                { "document.insertPageBreak", "F_insert_page_break" },
                { "document.listOperations", "F_list_document_operations" },
                { "document.restoreCheckpoint", "F_restore_document_checkpoint" },
                { "table.createFromXml", "F_create_table_from_xml" },
                { "table.applyFormatFromXml", "F_apply_table_format_from_xml" },
                { "table.extractFormat", "F_extract_table_format" },
                { "table.optimizeDisplay", "F_optimize_table_display" },
                { "table.delete", "F_delete_table" },
                { "table.applySourceFormat", "F_apply_source_table_format" },
                { "chart.createFromXml", "F_create_chart_from_xml" },
                { "chart.extractXml", "F_extract_chart_xml" },
                { "chart.applyFromXml", "F_apply_chart_from_xml" },
                { "chart.delete", "F_delete_chart" },
                { "document.insertSvgImage", "F_insert_svg_image" },
                { "document.insertImage", "F_insert_image" },
                { "document.deleteImage", "F_delete_image" },
                { "file.write", "F_write_file" },
                { "file.read", "F_read_file" },
                { "workspace.info", "F_workspace_info" },
                { "workspace.copyFile", "F_copy_workspace_file" },
                { "file.modifyYaml", "F_modify_yaml_file" },
                { "file.csvToXml", "F_csv_to_xml" },
                { "misc.greet", "F_greet" },
                { "document.getFormatContext", "F_get_format_context" },
                { "document.applyDocumentFormat", "F_apply_document_format" },
                { "document.applyParagraphFormat", "F_apply_paragraph_format" },
                { "document.getFormatDisplay", "F_get_format_display" },
                { "document.getTableFormatDisplay", "F_get_table_format_display" },
                // { "test.wordDocumentExtractor", "F_test_word_document_extractor" },
                { "test.paragraphCodeLocator", "F_test_paragraph_code_locator" },
            };
        }

        public IList<string> GetMethodNames()
        {
            var names = _methodToToolName.Keys.ToList();
            names.Add("workspace.pullFromBackend");
            names.Add("document.writeCaptureImage");
            return names;
        }

        public Task<ToolResult> InvokeToolNameAsync(string toolName, Dictionary<string, object> parameters)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                return Task.FromResult(new ToolResult { Success = false, Error = "tool is required" });
            }

            foreach (KeyValuePair<string, string> kv in _methodToToolName)
            {
                if (string.Equals(kv.Value, toolName, StringComparison.Ordinal))
                {
                    return InvokeRawAsync(kv.Key, parameters);
                }
            }

            return Task.FromResult(new ToolResult
            {
                Success = false,
                Error = "unknown tool '" + toolName + "'"
            });
        }

        public async Task<ToolResult> InvokeRawAsync(string method, Dictionary<string, object> parameters)
        {
            if (string.IsNullOrWhiteSpace(method))
            {
                return new ToolResult { Success = false, Error = "method is required" };
            }

            if (string.Equals(method, "workspace.pullFromBackend", StringComparison.Ordinal))
            {
                string conversationId = null;
                if (parameters != null && parameters.ContainsKey("conversation_id"))
                {
                    conversationId = parameters["conversation_id"]?.ToString();
                }

                return await WorkspaceReconcile.PullAssumingLockedAsync(conversationId).ConfigureAwait(false);
            }

            if (string.Equals(method, "document.writeCaptureImage", StringComparison.Ordinal))
            {
                return await WriteCaptureImageHelper.ExecuteAsync(parameters ?? new Dictionary<string, object>())
                    .ConfigureAwait(false);
            }

            if (!_methodToToolName.TryGetValue(method, out var toolName))
            {
                if (!_toolHandlers.ContainsKey(method))
                {
                    return new ToolResult { Success = false, Error = $"unknown method '{method}'" };
                }

                toolName = method;
            }

            if (!_toolHandlers.TryGetValue(toolName, out var handler))
            {
                return new ToolResult { Success = false, Error = $"handler not registered for method '{method}'" };
            }

            if (ToolDrainState.IsMutatingActive)
            {
                ToolStopRiskKind incoming = ConfigManager.ClassifyToolStopRisk(toolName);
                if (incoming == ToolStopRiskKind.Low || incoming == ToolStopRiskKind.High)
                {
                    await ToolDrainState.WaitIdleAsync().ConfigureAwait(false);
                }
            }

            var args = parameters ?? new Dictionary<string, object>();
            bool operateDance = ForegroundDanceTriggers.Classify(toolName, args)
                == ForegroundDanceKind.Operate;

            if (operateDance)
            {
                HostCallbacks.RaiseForegroundDance(new ForegroundDanceRequest
                {
                    Kind = ForegroundDanceKind.Operate,
                    ChannelId = ChannelContext.TryGetChannelIdFromParameters(args),
                    TargetHwnd = ForegroundDanceHwnd.TryResolveFromArgs(args),
                    Source = toolName,
                });
            }

            try
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }

                ToolResult result;
                if (WorkspaceReconcile.ShouldWrapFrontendWrite(toolName, args))
                {
                    int ttl = WorkspaceReconcile.TtlSecForTool(toolName, args);
                    result = await WorkspaceReconcile
                        .AroundFrontendWriteAsync(ttl, () => InvokeHandlerCoreAsync(toolName, handler, args))
                        .ConfigureAwait(false);
                }
                else
                {
                    result = await InvokeHandlerCoreAsync(toolName, handler, args).ConfigureAwait(false);
                }

                if (AgentRunCancellation.IsCancelled)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }

                return result;
            }
            finally
            {
                if (operateDance)
                {
                    HostCallbacks.RaiseForegroundDanceComplete();
                }
            }
        }

        private async Task<ToolResult> InvokeHandlerCoreAsync(
            string toolName,
            Func<Dictionary<string, object>, Task<ToolResult>> handler,
            Dictionary<string, object> args)
        {
            if (toolName == "F_list_document_operations" || toolName == "F_restore_document_checkpoint")
            {
                return await handler(args);
            }

            if (!DocumentCheckpointMutatingTools.ShouldCheckpoint(toolName, args))
            {
                return await handler(args);
            }

            if (DocumentCheckpointMutatingTools.TryGetContractParseFailure(
                    toolName, args, out ToolResult parseFail))
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "invoke-skip-ckpt-parse",
                    parseFail == null ? "" : parseFail.Error);
                return parseFail;
            }

            if (TryReturnBeforeCheckpoint(toolName, args, out ToolResult early))
            {
                return early;
            }

            CheckpointBeforeResult before;
            using (EasyWriteDiagnostics.Time("checkpoint.CreateBeforeMutation"))
            {
                before = DocumentCheckpointService.CreateBeforeMutation(toolName, args);
            }
            if (!before.Success)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "invoke-ckpt-fail",
                    "tool=" + toolName + " err=" + (before.Error ?? ""));
                return new ToolResult { Success = false, Error = before.Error };
            }

            WordAddIn1.DocumentHost.HostProbe.Log("invoke-ckpt-ok", "tool=" + toolName);

            ToolResult result;
            try
            {
                result = await handler(args);
            }
            catch (Exception)
            {
                try
                {
                    DocumentCheckpointService.RollbackFailedMutation(
                        before.DocUuid, before.Timestamp, before.ChannelId);
                }
                catch (Exception rollbackEx)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "invoke-rollback-threw",
                        WordAddIn1.DocumentHost.HostProbe.Ex(rollbackEx));
                }

                throw;
            }

            if (DocumentCheckpointService.IsInvokeMutationFailed(result, toolName))
            {
                try
                {
                    if (DocumentCheckpointService.IsCheckpointCleanupOnly(result))
                    {
                        DocumentCheckpointService.CleanupFailedInvoke(before.DocUuid, before.Timestamp);
                    }
                    else
                    {
                        DocumentCheckpointService.RollbackFailedMutation(
                            before.DocUuid, before.Timestamp, before.ChannelId);
                    }
                }
                catch (Exception ex)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log("invoke-rollback-threw", WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                }

                return result;
            }

            DocumentCheckpointService.AppendActionLogAfterMutation(toolName, args);
            return result;
        }

        /// <summary>
        /// 清除 / 写入在建副本之前探测引用。没有引用的清除直接成功；已有引用的写入直接失败。
        /// </summary>
        private bool TryReturnBeforeCheckpoint(
            string toolName,
            Dictionary<string, object> args,
            out ToolResult early)
        {
            early = null;
            bool isClear = toolName == "F_clear_data_provenance";
            bool isApply = toolName == "F_apply_data_provenance";
            if (!isClear && !isApply)
            {
                return false;
            }

            if (!DocumentHost.DocumentHostAdapter.TryResolveInteropDocument(
                    args,
                    _wordApplication,
                    out DocumentHost.InteropDocumentHandle handle,
                    out ToolResult resolveError))
            {
                early = resolveError ?? new ToolResult { Success = false, Error = "无法解析文档渠道" };
                return true;
            }

            bool has;
            try
            {
                has = DataProvenanceClearHelper.HasProvenance(handle.Document);
            }
            catch (Exception ex)
            {
                early = new ToolResult { Success = false, Error = "执行失败: " + ex.Message };
                return true;
            }

            if (isClear && !has)
            {
                early = new ToolResult
                {
                    Success = true,
                    Data = new { cleared = false },
                };
                return true;
            }

            if (isApply && has)
            {
                early = new ToolResult
                {
                    Success = false,
                    Error = DataProvenanceClearHelper.AlreadyHasProvenanceError,
                };
                return true;
            }

            return false;
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
