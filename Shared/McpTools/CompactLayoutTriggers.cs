using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// Desktop 缩小版自动触发白名单（与 DocumentCheckpointMutatingTools 独立）。
    /// 写文档/表格/演示的突变工具进名单；同一工具上的只读 action（list/get/extract）不缩。
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
            "F_ppt_transition",
        };

        /// <summary>只读 action：不触发缩小版（即使用具名在白名单内）。</summary>
        private static readonly HashSet<string> ReadOnlyActions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "list",
            "get",
            "extract",
        };

        public static bool ShouldCompact(string toolName)
        {
            return ShouldCompact(toolName, null);
        }

        public static bool ShouldCompact(string toolName, IReadOnlyDictionary<string, object> args)
        {
            if (string.IsNullOrEmpty(toolName) || !ToolNames.Contains(toolName))
            {
                return false;
            }

            string action = TryGetAction(args);
            if (!string.IsNullOrEmpty(action) && ReadOnlyActions.Contains(action))
            {
                return false;
            }

            return true;
        }

        private static string TryGetAction(IReadOnlyDictionary<string, object> args)
        {
            if (args == null)
            {
                return "";
            }

            if (args.TryGetValue("action", out object raw) && raw != null)
            {
                return Convert.ToString(raw)?.Trim() ?? "";
            }

            foreach (KeyValuePair<string, object> kv in args)
            {
                if (string.Equals(kv.Key, "action", StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    return Convert.ToString(kv.Value)?.Trim() ?? "";
                }
            }

            return "";
        }
    }
}
