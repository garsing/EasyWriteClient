using System;
using System.Collections.Generic;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class SentenceCodeLocator
    {
        public static Word.Range LocateRange(
            Word.Document doc,
            string code,
            int occurrenceIndex,
            string explicitTableId,
            TableScopeIndex tableScope,
            bool allowAutoTableScope = false,
            string debugTag = null)
        {
            string content = DocumentState.GetSentenceContent(code);
            if (string.IsNullOrEmpty(content))
            {
                return null;
            }

            string tableId = explicitTableId;
            if (string.IsNullOrEmpty(tableId) && allowAutoTableScope)
            {
                tableId = TableSentenceScopeHelper.ResolveEffectiveTableId(code, null, tableScope);
            }

            string tag = string.IsNullOrEmpty(debugTag) ? code : debugTag;
            return StoredTextRangeLocator.Locate(doc, content, occurrenceIndex, tableId, tag);
        }

        public static Word.Range LocateRangeByDisplayStart(
            Word.Document doc,
            string code,
            int displayStartPosition,
            string explicitTableId,
            TableScopeIndex tableScope,
            string debugTag = null)
        {
            return DisplayPositionRangeResolver.ResolveSingle(
                doc,
                code,
                displayStartPosition,
                explicitTableId,
                tableScope,
                domainOrderedCodes: null,
                debugTag: string.IsNullOrEmpty(debugTag) ? code : debugTag);
        }

        public static Word.Table ResolveTableById(Word.Document doc, string tableId)
        {
            if (doc == null || string.IsNullOrEmpty(tableId))
            {
                return null;
            }

            int tableOrderIndex = DocumentState.GetTableIndex(tableId);
            if (tableOrderIndex < 0)
            {
                return null;
            }

            List<Word.Table> allTables = GetAllTablesInOrder(doc);
            if (tableOrderIndex >= allTables.Count)
            {
                return null;
            }

            return allTables[tableOrderIndex];
        }

        public static List<Word.Table> GetAllTablesInOrder(Word.Document document)
        {
            var allTables = new List<Word.Table>();

            try
            {
                foreach (Word.Table table in document.Tables)
                {
                    CollectTablesRecursive(table, allTables);
                }

                allTables.Sort((t1, t2) => t1.Range.Start.CompareTo(t2.Range.Start));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SentenceLocator] 收集表格失败：{ex.Message}");
            }

            return allTables;
        }

        private static void CollectTablesRecursive(Word.Table table, List<Word.Table> allTables)
        {
            if (table == null)
            {
                return;
            }

            try
            {
                allTables.Add(table);

                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        foreach (Word.Table nestedTable in cell.Tables)
                        {
                            CollectTablesRecursive(nestedTable, allTables);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SentenceLocator] 递归收集表格失败：{ex.Message}");
            }
        }
    }
}
