using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// Desktop 缩小版自动触发白名单（与 DocumentCheckpointMutatingTools 独立）。
    /// </summary>
    public static class CompactLayoutTriggers
    {
        private static readonly HashSet<string> ToolNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "F_process_document_actions",
            "F_insert_document_from_file",
            "F_replace_document_from_file",
            "F_apply_document_format",
            "F_apply_paragraph_format",
            "F_apply_kb_format",
            "F_apply_kb_paragraph_format",
            "F_format_transfer",
            "F_apply_page_setup",
            "F_create_table_from_xml",
            "F_apply_table_format_from_xml",
            "F_apply_table_row_values",
            "F_optimize_table_display",
            "F_delete_table",
            "F_apply_kb_table_format",
            "F_create_chart_from_xml",
            "F_apply_chart_from_xml",
            "F_delete_chart",
            "F_insert_svg_image",
            "F_insert_workspace_image",
            "F_apply_data_provenance",
            "F_restore_document_checkpoint",
            "F_write_excel_range",
            "F_manage_excel_sheet",
            "F_apply_excel_format",
            "F_apply_excel_conditional_format",
            "F_apply_excel_structure",
            "F_excel_pivot",
            "F_excel_chart",
            "F_apply_ppt_html",
            "F_manage_ppt_slide",
            "F_ppt_animation",
        };

        public static bool ShouldCompact(string toolName)
        {
            return !string.IsNullOrEmpty(toolName) && ToolNames.Contains(toolName);
        }
    }
}
