using System;
using System.Collections.Generic;

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
            "F_apply_source_format",
            "F_apply_source_paragraph_format",
            "F_apply_source_table_format",
            "F_format_transfer",
            "F_apply_page_setup",
            "F_create_table_from_xml",
            "F_apply_table_format_from_xml",
            "F_apply_table_data_from_xml",
            "F_apply_table_row_values",
            "F_optimize_table_display",
            "F_delete_table",
            "F_create_chart_from_xml",
            "F_apply_chart_from_xml",
            "F_delete_chart",
            "F_insert_svg_image",
            "F_insert_image",
            "F_apply_data_provenance",
        };

        public static bool ShouldCheckpoint(string toolName, IReadOnlyDictionary<string, object> parameters)
        {
            if (string.IsNullOrEmpty(toolName) || !MutatingToolNames.Contains(toolName))
            {
                return false;
            }

            return true;
        }
    }
}
