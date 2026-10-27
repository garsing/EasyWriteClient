/**
 * 连续工具折叠的探索名单。与需求 §3.1 同一份。
 * 不在名单里、但会画出工具卡的，组标题算「操作」。
 * 不复用 ToolCallBoxDisplayRule 的读/改图标名单。
 */
export const EXPLORE_TOOL_NAMES = [
  'B_analyze_kb_document_page_image',
  'B_calculate',
  'B_capture_kb_document_page_image',
  'B_data_fetcher',
  'B_get_kb_chunk_content',
  'B_get_kb_document_page_info',
  'B_get_kb_format_display',
  'B_get_kb_table_format_display',
  'B_interface_doc',
  'B_list_kb_document_chunks',
  'B_read_pdf',
  'B_read_skill_file',
  'B_search_conversation_memory',
  'B_search_knowledge_base',
  'F_analyze_document_page_image',
  'F_browser_snapshot',
  'F_capture_document_page_image',
  'F_capture_ppt_shape',
  'F_extract_chart_xml',
  'F_extract_table_data',
  'F_extract_table_format',
  'F_find_position',
  'F_get_curr_doc_chunk_display_content',
  'F_get_curr_doc_chunk_paragraph_display_content',
  'F_get_data_provenance',
  'F_get_document_content',
  'F_get_document_page_info',
  'F_get_document_structure',
  'F_get_excel_format',
  'F_get_format_context',
  'F_get_format_display',
  'F_get_presentation_content',
  'F_get_table_format_display',
  'F_get_workbook_content',
  'F_greet',
  'F_list_comments',
  'F_list_document_operations',
  'F_read_excel_range',
  'F_read_file',
  'F_read_ppt_html',
  'F_read_table_row_values',
  'F_test_heading_recognizer',
  'F_test_paragraph_code_locator',
  'F_test_word_document_extractor',
  'F_workspace_info'
]

const exploreSet = new Set(EXPLORE_TOOL_NAMES)

/**
 * @param {string} toolName
 * @returns {'explore'|'operate'|null} 没名字时 null，标题沿用上一张
 */
export function categoryOf (toolName) {
  const name = String(toolName || '').trim()
  if (!name) return null
  return exploreSet.has(name) ? 'explore' : 'operate'
}
