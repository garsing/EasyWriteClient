using System;
using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// P1：通过 DocumentState 持久化 S_/P_ 映射。
    /// </summary>
    public sealed class DocumentStateMappingStore : IMappingStore
    {
        public void BeginMappingPass()
        {
            DocumentState.ClearSentenceParagraphIndex();
        }

        public Tuple<string, bool> FindOrCreateSentence(string content, int mark)
        {
            string name = DocumentState.FindOrCreateSentenceName(content, mark);
            return Tuple.Create(name, false);
        }

        public Tuple<string, bool> FindOrCreateParagraph(string storedText, int mark)
        {
            string name = DocumentState.FindOrCreateParagraphName(storedText, mark);
            return Tuple.Create(name, false);
        }

        public void LinkSentenceToParagraph(string sentenceName, string paragraphName)
        {
            DocumentState.LinkSentenceToParagraph(sentenceName, paragraphName);
        }

        public IReadOnlyList<IReadOnlyList<string>> GetSentenceMappings() =>
            DocumentState.SentenceNameMapping;

        public IReadOnlyList<IReadOnlyList<string>> GetParagraphMappings() =>
            DocumentState.ParagraphNameMapping;

        public IReadOnlyDictionary<string, string> GetSentenceToParagraph() =>
            DocumentState.SentenceToParagraph;

        public IReadOnlyDictionary<string, List<string>> GetParagraphToSentences() =>
            DocumentState.ParagraphToSentences;
    }
}
