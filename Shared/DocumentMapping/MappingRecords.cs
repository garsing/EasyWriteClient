using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L2：按 \r 切分后的段落块。
    /// </summary>
    public sealed class ParagraphBlock
    {
        public int Index { get; }
        public string Text { get; }
        public bool TrailingCr { get; }

        public ParagraphBlock(int index, string text, bool trailingCr)
        {
            Index = index;
            Text = text ?? "";
            TrailingCr = trailingCr;
        }

        public string StoredText => Text + (TrailingCr ? "\r" : "");
    }

    public sealed class SegMapResult
    {
        public List<string> ParagraphNames { get; } = new List<string>();
        public List<string> SentenceNames { get; } = new List<string>();
    }

    public sealed class MapReadTextResult
    {
        public List<List<string>> SentenceMappings { get; set; } = new List<List<string>>();
        public List<List<string>> ParagraphMappings { get; set; } = new List<List<string>>();
        public Dictionary<string, string> SentenceToParagraph { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, List<string>> ParagraphToSentences { get; set; } = new Dictionary<string, List<string>>();
        public List<string> ChunkDisplayContents { get; set; } = new List<string>();
        public List<string> ChunkParagraphDisplayContents { get; set; } = new List<string>();
        public List<List<string>> AllChunkParagraphNames { get; set; } = new List<List<string>>();
        public List<List<string>> AllChunkSentenceNames { get; set; } = new List<List<string>>();
        public List<List<string>> AllChunkNonSegs { get; set; } = new List<List<string>>();
        public List<List<List<string>>> AllChunkSegsSentenceNames { get; set; } = new List<List<List<string>>>();
        public List<List<List<string>>> AllChunkSegsParagraphNames { get; set; } = new List<List<List<string>>>();
        public List<List<int>> AllChunkSegMarks { get; set; } = new List<List<int>>();
        public Dictionary<string, string> NameToSentenceContentMap { get; set; } = new Dictionary<string, string>();
    }
}
