using System;
using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L3 映射存储端口（DocumentState 或内存实现）。
    /// </summary>
    public interface IMappingStore
    {
        void BeginMappingPass();

        Tuple<string, bool> FindOrCreateSentence(string content, int mark);

        Tuple<string, bool> FindOrCreateParagraph(string storedText, int mark);

        void LinkSentenceToParagraph(string sentenceName, string paragraphName);

        IReadOnlyList<IReadOnlyList<string>> GetSentenceMappings();

        IReadOnlyList<IReadOnlyList<string>> GetParagraphMappings();

        IReadOnlyDictionary<string, string> GetSentenceToParagraph();

        IReadOnlyDictionary<string, List<string>> GetParagraphToSentences();
    }
}
