using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class SentenceCodeLocator
    {
        public static Word.Table ResolveTableById(Word.Document doc, string tableId)
        {
            return TableResolveHelper.TryResolveTableById(doc, tableId, out _);
        }

        public static List<Word.Table> GetAllTablesInOrder(Word.Document document)
        {
            return TableResolveHelper.GetAllTablesInOrder(document);
        }
    }
}
