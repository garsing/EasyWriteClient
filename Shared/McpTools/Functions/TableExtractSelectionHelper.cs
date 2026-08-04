using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class TableExtractSelectionResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public Word.Table Table { get; set; }
        public int TableIndex { get; set; }
        public string TableId { get; set; }
        public string SelectionType { get; set; }
    }

    public static class TableExtractSelectionHelper
    {
        public static TableExtractSelectionResult Resolve(Word.Document document, Dictionary<string, object> tableSelection)
        {
            var result = new TableExtractSelectionResult();
            if (document == null)
            {
                result.Error = "没有活动的Word文档";
                return result;
            }

            if (tableSelection == null)
            {
                result.Error = "table_selection必须是对象格式";
                return result;
            }

            string selectionType = tableSelection.ContainsKey("type")
                ? tableSelection["type"]?.ToString()
                : "";

            if (selectionType == "by_id")
            {
                return ResolveById(document, tableSelection, selectionType);
            }

            if (selectionType == "by_context")
            {
                result.Error =
                    "选择类型 by_context 暂不支持，请使用 by_id 并提供 table_id（可先 get_document_content 获取 T_ 编号）";
                return result;
            }

            result.Error = $"不支持的选择类型：{selectionType}，请使用 by_id";
            return result;
        }

        private static TableExtractSelectionResult ResolveById(
            Word.Document document,
            Dictionary<string, object> tableSelection,
            string selectionType)
        {
            var result = new TableExtractSelectionResult { SelectionType = selectionType };

            if (!tableSelection.ContainsKey("table_id"))
            {
                result.Error = "选择类型为by_id时，必须提供table_id参数";
                return result;
            }

            string tableId = tableSelection["table_id"]?.ToString();
            if (string.IsNullOrEmpty(tableId))
            {
                result.Error = "table_id不能为空";
                return result;
            }

            int tableOrderIndex = DocumentState.GetTableIndex(tableId);
            if (tableOrderIndex < 0)
            {
                result.Error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                return result;
            }

            List<Word.Table> allTables = SentenceCodeLocator.GetAllTablesInOrder(document);
            if (tableOrderIndex >= allTables.Count)
            {
                result.Error =
                    $"表格编号 '{tableId}' 对应的索引{tableOrderIndex}无效，文档中只有{allTables.Count}个表格（包括嵌套表格）";
                return result;
            }

            try
            {
                result.Table = allTables[tableOrderIndex];
                result.TableIndex = GetTableIndex(document, result.Table);
                result.TableId = tableId;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Error = $"获取表格（编号：{tableId}）失败：{ex.Message}";
                return result;
            }
        }

        private static TableExtractSelectionResult ResolveByContext(
            Word.Document document,
            Dictionary<string, object> tableSelection,
            string selectionType)
        {
            var result = new TableExtractSelectionResult { SelectionType = selectionType };

            string contextBefore = tableSelection.ContainsKey("context_before")
                ? tableSelection["context_before"]?.ToString() ?? ""
                : "";
            string contextAfter = tableSelection.ContainsKey("context_after")
                ? tableSelection["context_after"]?.ToString() ?? ""
                : "";

            if (string.IsNullOrEmpty(contextBefore))
            {
                result.Error = "选择类型为by_context时，必须提供context_before参数";
                return result;
            }

            try
            {
                Word.Range beforeRange = McpToolsHelpers.FindText(document.Content, contextBefore);
                if (beforeRange == null)
                {
                    result.Error = $"找不到上下文'{contextBefore}'";
                    return result;
                }

                foreach (Word.Table table in document.Tables)
                {
                    if (table.Range.Start < beforeRange.End)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(contextAfter))
                    {
                        Word.Range afterRange = McpToolsHelpers.FindText(document.Content, contextAfter);
                        if (afterRange != null && table.Range.End <= afterRange.Start)
                        {
                            result.Table = table;
                            result.TableIndex = GetTableIndex(document, table);
                            result.Success = true;
                            return result;
                        }
                    }
                    else
                    {
                        result.Table = table;
                        result.TableIndex = GetTableIndex(document, table);
                        result.Success = true;
                        return result;
                    }
                }

                result.Error = "找不到上下文匹配的表格";
                return result;
            }
            catch (Exception ex)
            {
                result.Error = $"通过上下文查找表格失败：{ex.Message}";
                return result;
            }
        }

        private static int GetTableIndex(Word.Document document, Word.Table table)
        {
            try
            {
                for (int i = 1; i <= document.Tables.Count; i++)
                {
                    if (document.Tables[i].Range.Start == table.Range.Start)
                    {
                        return i;
                    }
                }
            }
            catch
            {
            }

            return 1;
        }
    }
}
