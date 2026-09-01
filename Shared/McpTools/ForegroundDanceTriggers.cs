using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// 桌面前台编排分类：打开成功后 / 操作前 / 不编排。
    /// 不再用 CompactLayoutTriggers 决定「该不该缩」。
    /// </summary>
    public static class ForegroundDanceTriggers
    {
        private static readonly HashSet<string> OpenToolNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "F_open_document",
            "F_open_word_document",
            "F_browser_navigate",
        };

        /// <summary>
        /// 原 Compact 写名单去掉 navigate，加上 interact。
        /// </summary>
        private static readonly HashSet<string> OperateToolNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "F_process_document_actions",
            "F_insert_document_from_file",
            "F_replace_document_from_file",
            "F_apply_document_format",
            "F_apply_paragraph_format",
            "F_apply_source_format",
            "F_apply_source_paragraph_format",
            "F_format_transfer",
            "F_apply_page_setup",
            "F_create_table_from_xml",
            "F_apply_table_format_from_xml",
            "F_apply_table_row_values",
            "F_optimize_table_display",
            "F_delete_table",
            "F_apply_source_table_format",
            "F_create_chart_from_xml",
            "F_apply_chart_from_xml",
            "F_delete_chart",
            "F_insert_svg_image",
            "F_insert_image",
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
            "F_manage_ppt_shape",
            "F_ppt_animation",
            "F_ppt_transition",
            "F_browser_interact",
        };

        private static readonly HashSet<string> ReadOnlyActions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "list",
            "get",
            "extract",
        };

        public static ForegroundDanceKind Classify(string toolName, IReadOnlyDictionary<string, object> args)
        {
            if (string.IsNullOrEmpty(toolName))
            {
                return ForegroundDanceKind.None;
            }

            if (OpenToolNames.Contains(toolName))
            {
                if (string.Equals(toolName, "F_browser_navigate", StringComparison.Ordinal)
                    && !CompactLayoutTriggers.IsBrowserNavigateVisible(args))
                {
                    return ForegroundDanceKind.None;
                }

                return ForegroundDanceKind.Open;
            }

            if (!OperateToolNames.Contains(toolName))
            {
                return ForegroundDanceKind.None;
            }

            string action = TryGetAction(args);
            if (!string.IsNullOrEmpty(action) && ReadOnlyActions.Contains(action))
            {
                return ForegroundDanceKind.None;
            }

            return ForegroundDanceKind.Operate;
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
