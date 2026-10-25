using System;
using System.Collections.Generic;
using System.IO;

namespace WordAddIn1
{
    public static class DocumentCheckpointMutatingTools
    {
        private static readonly HashSet<string> MutatingToolNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "F_process_document_actions",
            "F_insert_document_from_file",
            "F_replace_document_from_file",
            "F_apply_document_format",
            "F_apply_paragraph_format",
            "F_apply_source_format",
            "F_apply_source_paragraph_format",
            "F_apply_source_table_format",
            "F_set_track_revisions",
            "F_format_transfer",
            "F_apply_page_setup",
            "F_create_table_from_xml",
            "F_apply_table_format_from_xml",
            "F_apply_table_data_from_xml",
            "F_optimize_table_display",
            "F_delete_table",
            "F_create_chart_from_xml",
            "F_apply_chart_from_xml",
            "F_delete_chart",
            "F_insert_image",
            "F_delete_image",
            "F_insert_blank_page",
            "F_insert_page_break",
            "F_process_comments",
            "F_apply_data_provenance",
            "F_clear_data_provenance",
            "F_apply_ppt_html",
            "F_manage_ppt_shape",
            "F_manage_ppt_slide",
            "F_manage_ppt_lib",
            "F_ppt_transition",
            "F_ppt_animation",
            "F_write_excel_range",
            "F_apply_excel_format",
            "F_apply_excel_conditional_format",
            "F_apply_excel_structure",
            "F_excel_pivot",
            "F_excel_chart",
            "F_manage_excel_sheet",
            "F_excel_image",
        };

        public static bool ShouldCheckpoint(string toolName, IReadOnlyDictionary<string, object> parameters)
        {
            if (string.IsNullOrEmpty(toolName) || !MutatingToolNames.Contains(toolName))
            {
                return false;
            }

            if (string.Equals(toolName, "F_excel_image", StringComparison.Ordinal))
            {
                string action = null;
                if (parameters != null && parameters.TryGetValue("action", out object raw) && raw != null)
                {
                    action = Convert.ToString(raw);
                }

                if (string.IsNullOrWhiteSpace(action))
                {
                    return false;
                }

                action = action.Trim().ToLowerInvariant();
                return action == "insert" || action == "update" || action == "delete";
            }

            return true;
        }

        public static ToolResult Fail(string error)
        {
            return new ToolResult
            {
                Success = false,
                Error = error,
                Data = new Dictionary<string, object>
                {
                    ["checkpoint_cleanup_only"] = true
                }
            };
        }

        /// <summary>
        /// 不碰稿的契约。失败则不应建副本。无契约或通过返回 false。
        /// </summary>
        public static bool TryGetContractParseFailure(
            string toolName,
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            if (string.IsNullOrEmpty(toolName))
            {
                return false;
            }

            switch (toolName)
            {
                case "F_process_document_actions":
                    return F_ProcessDocumentActionsTool.TryGetContractParseFailure(args, out failure);
                case "F_insert_document_from_file":
                    return F_InsertDocumentFromFileTool.TryGetContractParseFailure(args, out failure);
                case "F_replace_document_from_file":
                    return F_ReplaceDocumentFromFileTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_document_format":
                    return F_ApplyDocumentFormatTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_source_table_format":
                    return F_ApplySourceTableFormatTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_page_setup":
                    return F_ApplyPageSetupTool.TryGetContractParseFailure(args, out failure);
                case "F_create_table_from_xml":
                    return F_CreateTableFromXmlTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_table_format_from_xml":
                    return F_ApplyTableFormatFromXmlTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_table_data_from_xml":
                    return F_ApplyTableDataFromXmlTool.TryGetContractParseFailure(args, out failure);
                case "F_optimize_table_display":
                    return F_OptimizeTableDisplayTool.TryGetContractParseFailure(args, out failure);
                case "F_delete_table":
                    return F_DeleteTableTool.TryGetContractParseFailure(args, out failure);
                case "F_create_chart_from_xml":
                    return F_CreateChartFromXmlTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_chart_from_xml":
                    return F_ApplyChartFromXmlTool.TryGetContractParseFailure(args, out failure);
                case "F_delete_chart":
                    return F_DeleteChartTool.TryGetContractParseFailure(args, out failure);
                case "F_insert_image":
                    return F_InsertWorkspaceImageTool.TryGetContractParseFailure(args, out failure);
                case "F_delete_image":
                    return F_DeleteImageTool.TryGetContractParseFailure(args, out failure);
                case "F_insert_blank_page":
                    return F_InsertBlankPageTool.TryGetContractParseFailure(args, out failure);
                case "F_insert_page_break":
                    return F_InsertPageBreakTool.TryGetContractParseFailure(args, out failure);
                case "F_process_comments":
                    return F_ProcessCommentsTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_data_provenance":
                    return F_ApplyDataProvenanceTool.TryGetContractParseFailure(args, out failure);
                case "F_clear_data_provenance":
                    return false;
                case "F_apply_ppt_html":
                    return F_ApplyPptHtmlTool.TryGetContractParseFailure(args, out failure);
                case "F_manage_ppt_shape":
                    return F_ManagePptShapeTool.TryGetContractParseFailure(args, out failure);
                case "F_manage_ppt_slide":
                    return F_ManagePptSlideTool.TryGetContractParseFailure(args, out failure);
                case "F_manage_ppt_lib":
                    return F_ManagePptLibTool.TryGetContractParseFailure(args, out failure);
                case "F_ppt_transition":
                    return F_PptTransitionTool.TryGetContractParseFailure(args, out failure);
                case "F_ppt_animation":
                    return F_PptAnimationTool.TryGetContractParseFailure(args, out failure);
                case "F_write_excel_range":
                    return F_WriteExcelRangeTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_excel_format":
                    return F_ApplyExcelFormatTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_excel_conditional_format":
                    return F_ApplyExcelConditionalFormatTool.TryGetContractParseFailure(args, out failure);
                case "F_apply_excel_structure":
                    return F_ApplyExcelStructureTool.TryGetContractParseFailure(args, out failure);
                case "F_excel_pivot":
                    return F_ExcelPivotTool.TryGetContractParseFailure(args, out failure);
                case "F_excel_chart":
                    return F_ExcelChartTool.TryGetContractParseFailure(args, out failure);
                case "F_excel_image":
                    return F_ExcelImageTool.TryGetContractParseFailure(args, out failure);
                case "F_manage_excel_sheet":
                    return F_ManageExcelSheetTool.TryGetContractParseFailure(args, out failure);
                default:
                    return false;
            }
        }

        public static bool TryReadLocalText(ResolvedFilePath resolved, out string text, out string error)
        {
            text = null;
            error = null;
            if (resolved == null || string.IsNullOrEmpty(resolved.LocalPath))
            {
                error = "路径未解析";
                return false;
            }

            if (!File.Exists(resolved.LocalPath))
            {
                error = "文件不存在: " + resolved.Display;
                return false;
            }

            try
            {
                text = File.ReadAllText(resolved.LocalPath);
            }
            catch (Exception ex)
            {
                error = "读取文件失败: " + ex.Message;
                return false;
            }

            if (string.IsNullOrEmpty(text))
            {
                error = "文件内容为空: " + resolved.Display;
                return false;
            }

            return true;
        }
    }
}
